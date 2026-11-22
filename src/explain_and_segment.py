"""
Açıklanabilir Yapay Zeka (Grad-CAM) ve SAM Tümör Segmentasyonu (src/explain_and_segment.py)
-------------------------------------------------------------------------------------------
1. Özel PyTorch Grad-CAM:
   - CNN (ResNet-50), ViT (vit_base_patch16_224) ve Swin Transformer (swin_tiny_patch4_window7_224)
     mimarilerinin tamamında forward/backward hook ile karar ısı haritası üretir.
2. Otomatik Prompt Çıkarımı:
   - Grad-CAM ısı haritasındaki yüksek aktivasyon bölgesi eşiklenip morfolojik olarak temizlenir
     ve SAM için sınır kutusu (bbox) + tepe noktası (point prompt) elde edilir.
3. SAM / MobileSAM Segmentasyonu (Ultralytics):
   - Çıkarılan prompt ile piksel hassasiyetinde tümör/lezyon maskesi üretir,
     kontur çizimi ve tümör alan oranı (%) hesaplar.
"""

from __future__ import annotations

import math
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

import cv2
import numpy as np
import torch
import torch.nn as nn

from src.models import get_gradcam_target_layer


class UnifiedGradCAM:
    """
    CNN, ViT ve Swin Transformer mimarilerinin hepsini destekleyen
    temiz, bağımsız PyTorch Grad-CAM implementasyonu.
    """

    def __init__(self, model: nn.Module, model_key: str):
        self.model = model
        self.model_key = model_key
        self.target_layer = get_gradcam_target_layer(model, model_key)
        self.activations: Optional[torch.Tensor] = None
        self.gradients: Optional[torch.Tensor] = None
        self._handles = [
            self.target_layer.register_forward_hook(self._save_activation),
            self.target_layer.register_full_backward_hook(self._save_gradient),
        ]

    def _save_activation(self, module: nn.Module, inp: Any, output: torch.Tensor) -> None:
        self.activations = output.detach()

    def _save_gradient(self, module: nn.Module, grad_input: Any, grad_output: Tuple[torch.Tensor, ...]) -> None:
        self.gradients = grad_output[0].detach()

    def remove_hooks(self) -> None:
        for h in self._handles:
            h.remove()
        self._handles.clear()

    def _reshape_to_bchw(self, tensor: torch.Tensor) -> torch.Tensor:
        """
        Katman çıktısını mimariye göre standart (B, C, H, W) biçimine dönüştürür:
        - CNN:  (B, C, H, W) -> doğrudan kullanılır
        - ViT:  (B, 1+N, C)  -> CLS token atılır, (B, C, sqrt(N), sqrt(N)) yapılır
        - Swin: (B, H, W, C) veya (B, N, C) -> (B, C, H, W) yapılır
        """
        if self.model_key == "cnn":
            return tensor

        if tensor.ndim == 4:
            # Swin timm çıktısı çoğunlukla (B, H, W, C)
            if tensor.shape[-1] >= tensor.shape[1] and tensor.shape[-1] >= tensor.shape[2]:
                return tensor.permute(0, 3, 1, 2).contiguous()
            return tensor

        if tensor.ndim == 3:
            b, n, c = tensor.shape
            if self.model_key == "vit":
                # CLS token (ilk token) çıkarılır
                tokens = tensor[:, 1:, :]
                n_spatial = tokens.shape[1]
            else:
                tokens = tensor
                n_spatial = n
            side = int(round(math.sqrt(n_spatial)))
            tokens = tokens[:, : side * side, :]
            return tokens.reshape(b, side, side, c).permute(0, 3, 1, 2).contiguous()

        raise ValueError(f"Beklenmeyen aktivasyon boyutu: {tuple(tensor.shape)}")

    def generate(
        self,
        input_tensor: torch.Tensor,
        target_class: Optional[int] = None,
        output_size: Tuple[int, int] = (224, 224),
    ) -> Tuple[np.ndarray, int, np.ndarray]:
        """
        Grad-CAM ısı haritası [0, 1], tahmin edilen sınıf indeksi ve olasılık dizisini döndürür.
        """
        self.model.zero_grad(set_to_none=True)
        with torch.enable_grad():
            logits = self.model(input_tensor)
            probs = torch.softmax(logits, dim=1)[0].detach().cpu().numpy()
            pred_idx = int(np.argmax(probs)) if target_class is None else int(target_class)

            score = logits[0, pred_idx]
            score.backward()

        assert self.activations is not None and self.gradients is not None, "Grad-CAM hook yakalanamadı."

        acts = self._reshape_to_bchw(self.activations).float()
        grads = self._reshape_to_bchw(self.gradients).float()

        # Kanal bazlı global ortalama ağırlık (alpha_k)
        weights = grads.mean(dim=(2, 3), keepdim=True)
        cam = (weights * acts).sum(dim=1, keepdim=False)[0]
        cam = torch.relu(cam).cpu().numpy()

        cam_min, cam_max = float(cam.min()), float(cam.max())
        if cam_max - cam_min > 1e-8:
            cam = (cam - cam_min) / (cam_max - cam_min)
        else:
            cam = np.zeros_like(cam)

        cam_resized = cv2.resize(cam, output_size, interpolation=cv2.INTER_CUBIC)
        cam_resized = np.clip(cam_resized, 0.0, 1.0)
        return cam_resized, pred_idx, probs


def overlay_gradcam(rgb_image: np.ndarray, cam_2d: np.ndarray, alpha: float = 0.45) -> np.ndarray:
    """Grad-CAM ısı haritasını RGB görüntü üzerine JET renk haritasıyla bindirir."""
    h, w = rgb_image.shape[:2]
    cam_rs = cv2.resize(cam_2d, (w, h), interpolation=cv2.INTER_CUBIC)
    heatmap_bgr = cv2.applyColorMap(np.uint8(255 * cam_rs), cv2.COLORMAP_JET)
    heatmap_rgb = cv2.cvtColor(heatmap_bgr, cv2.COLOR_BGR2RGB)
    blended = (1.0 - alpha) * rgb_image.astype(np.float32) + alpha * heatmap_rgb.astype(np.float32)
    return np.clip(blended, 0, 255).astype(np.uint8)


def compute_brain_interior_mask(rgb_image: np.ndarray, erode_px: int = 14) -> np.ndarray:
    """
    MR görüntüsünden siyah arka planı ve parlak kafatası çeperini ayırarak
    yalnızca beyin parankimi (iç bölge) ikili maskesini [0, 1] çıkarır.
    Transformer köşe-token (attention sink) artefaktlarını ve dış kenar taşmalarını önler.
    """
    gray = cv2.cvtColor(rgb_image, cv2.COLOR_RGB2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    _, mask = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    k_close = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (11, 11))
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, k_close, iterations=2)

    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    filled = np.zeros_like(mask)
    if contours:
        largest = max(contours, key=cv2.contourArea)
        cv2.drawContours(filled, [largest], -1, 255, thickness=cv2.FILLED)
    else:
        filled = np.full_like(mask, 255)

    # Kafatası parlak çeperini dışlamak için kontrollü morfolojik erozyon
    if erode_px > 0:
        k_erode = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (erode_px, erode_px))
        eroded = cv2.erode(filled, k_erode, iterations=1)
        if eroded.sum() > 0.15 * filled.sum():
            filled = eroded

    return (filled > 0).astype(np.uint8)


def mask_cam_to_brain(cam_2d: np.ndarray, rgb_image: np.ndarray, erode_px: int = 12) -> np.ndarray:
    """Grad-CAM ısı haritasını beyin iç maskesiyle sınırlayıp yeniden [0, 1] normalize eder."""
    h, w = rgb_image.shape[:2]
    cam_rs = cv2.resize(cam_2d, (w, h), interpolation=cv2.INTER_CUBIC)
    brain_mask = compute_brain_interior_mask(rgb_image, erode_px=erode_px).astype(np.float32)
    # Yumuşak geçişli maskeleme
    soft_mask = cv2.GaussianBlur(brain_mask, (9, 9), 0)
    masked = cam_rs * soft_mask
    cmin, cmax = float(masked.min()), float(masked.max())
    if cmax - cmin > 1e-8:
        masked = (masked - cmin) / (cmax - cmin)
    return np.clip(masked, 0.0, 1.0)


def extract_prompt_from_cam(
    cam_2d: np.ndarray,
    rgb_image: Optional[np.ndarray] = None,
    threshold: float = 0.55,
    margin: int = 6,
) -> Tuple[Tuple[int, int, int, int], Tuple[int, int]]:
    """
    Grad-CAM ısı haritasından (beyin iç bölgesine kısıtlanmış olarak) SAM için
    otomatik sınır kutusu (x1, y1, x2, y2) ve tepe noktasını (px, py) çıkarır.
    """
    h, w = cam_2d.shape[:2]
    work_cam = mask_cam_to_brain(cam_2d, rgb_image, erode_px=12) if rgb_image is not None else cam_2d

    peak_flat = int(np.argmax(work_cam))
    py, px = divmod(peak_flat, w)

    binary = (work_cam >= threshold).astype(np.uint8) * 255
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (5, 5))
    binary = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel)
    binary = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel)

    contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        half = min(h, w) // 7
        x1, y1 = max(0, px - half), max(0, py - half)
        x2, y2 = min(w - 1, px + half), min(h - 1, py + half)
        return (x1, y1, x2, y2), (px, py)

    # Tepe noktasını içeren konturu, yoksa en büyük konturu seç
    chosen = max(contours, key=cv2.contourArea)
    for cnt in contours:
        if cv2.pointPolygonTest(cnt, (float(px), float(py)), False) >= 0:
            chosen = cnt
            break

    x, y, bw, bh = cv2.boundingRect(chosen)
    x1 = max(0, x - margin)
    y1 = max(0, y - margin)
    x2 = min(w - 1, x + bw + margin)
    y2 = min(h - 1, y + bh + margin)
    return (x1, y1, x2, y2), (px, py)


class TumorSegmenter:
    """
    Ultralytics MobileSAM tabanlı tümör segmentasyon motoru.
    Grad-CAM'den gelen kutu/nokta prompt'u (veya kullanıcının manuel kutusu) ile
    piksel hassasiyetinde maske ve tümör alanı oranı üretir.
    """

    def __init__(self, weights_name: str = "mobile_sam.pt", weights_dir: Optional[Path] = None):
        self.weights_dir = Path(weights_dir) if weights_dir else Path(__file__).resolve().parent.parent / "checkpoints"
        self.weights_dir.mkdir(parents=True, exist_ok=True)
        self.weights_path = self.weights_dir / weights_name
        self.sam_model = None
        self.backend_name = "MobileSAM (Ultralytics)"

    def _ensure_loaded(self) -> bool:
        if self.sam_model is not None:
            return True
        try:
            from ultralytics import SAM

            self.sam_model = SAM(str(self.weights_path))
            self.backend_name = "MobileSAM (Ultralytics SAM)"
            return True
        except Exception as exc:
            self.backend_name = f"Fallback (Otsu/CAM - {exc})"
            self.sam_model = None
            return False

    def segment(
        self,
        rgb_image: np.ndarray,
        cam_2d: np.ndarray,
        cam_threshold: float = 0.55,
        custom_bbox: Optional[Tuple[int, int, int, int]] = None,
        apply_interior_mask: bool = True,
    ) -> Dict[str, Any]:
        """
        Görüntü üzerinde tümör/lezyon segmentasyonu gerçekleştirir.
        - apply_interior_mask=True: MR ve BT gibi siyah dış çerçeveli görüntülerde
          kafatası/gövde dışını ve köşe artefaktlarını temizler.
        - apply_interior_mask=False: Göğüs Röntgeni gibi tüm çerçeveyi dolduran görüntülerde
          akciğer alanlarının korunmasını sağlar.
        """
        h, w = rgb_image.shape[:2]
        if apply_interior_mask:
            brain_interior = compute_brain_interior_mask(rgb_image, erode_px=10)
            cam_rs = mask_cam_to_brain(cam_2d, rgb_image, erode_px=12)
            auto_bbox, peak_point = extract_prompt_from_cam(cam_rs, rgb_image=rgb_image, threshold=cam_threshold)
        else:
            brain_interior = np.ones((h, w), dtype=np.uint8)
            cam_rs = cv2.resize(cam_2d, (w, h), interpolation=cv2.INTER_CUBIC)
            auto_bbox, peak_point = extract_prompt_from_cam(cam_rs, rgb_image=None, threshold=cam_threshold)
        bbox = custom_bbox if custom_bbox is not None else auto_bbox

        mask_bin: Optional[np.ndarray] = None
        method_used = self.backend_name

        # Negatif arka plan noktaları (köşeler) — SAM'in siyah arka planı seçmesini engeller
        neg_points = [[4, 4], [w - 5, 4], [4, h - 5], [w - 5, h - 5]]
        all_points = [[peak_point[0], peak_point[1]]] + neg_points
        all_labels = [1, 0, 0, 0, 0]

        if self._ensure_loaded():
            try:
                results = self.sam_model.predict(
                    source=rgb_image,
                    bboxes=[list(bbox)],
                    points=[all_points],
                    labels=[all_labels],
                    verbose=False,
                )
                if results and results[0].masks is not None and len(results[0].masks.data) > 0:
                    raw_mask = results[0].masks.data[0].detach().cpu().numpy()
                    raw_mask = cv2.resize(raw_mask.astype(np.float32), (w, h), interpolation=cv2.INTER_NEAREST)
                    candidate = ((raw_mask > 0.5).astype(np.uint8)) * brain_interior
                    # Maske tüm beynin %45'inden büyükse (aşırı taşma) bbox içine kırp
                    if candidate.sum() > 0.45 * max(brain_interior.sum(), 1):
                        bx1, by1, bx2, by2 = bbox
                        box_mask = np.zeros_like(candidate)
                        box_mask[by1:by2, bx1:bx2] = 1
                        candidate = candidate * box_mask
                    mask_bin = candidate
                    method_used = self.backend_name
            except Exception as exc:
                method_used = f"CAM-Morphology Fallback ({exc})"

        # Güvenli yedek (SAM maske üretemezse)
        if mask_bin is None or mask_bin.sum() == 0:
            mask_bin = ((cam_rs >= cam_threshold).astype(np.uint8)) * brain_interior
            kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (5, 5))
            mask_bin = cv2.morphologyEx(mask_bin, cv2.MORPH_CLOSE, kernel)
            if "SAM" in method_used:
                method_used = "Grad-CAM Morfolojik Yedek"

        # Beyin piksel alanı (siyah arka plan hariç)
        gray = cv2.cvtColor(rgb_image, cv2.COLOR_RGB2GRAY)
        brain_pixels = int((gray > 15).sum())
        tumor_pixels = int(mask_bin.sum())
        tumor_ratio_pct = (tumor_pixels / max(brain_pixels, 1)) * 100.0

        # Görselleştirme bindirmesi (Yarı saydam kırmızı dolgu + neon yeşil kontur + prompt kutusu)
        overlay = rgb_image.copy()
        red_layer = np.zeros_like(overlay)
        red_layer[:, :, 0] = 235
        red_layer[:, :, 1] = 45
        red_layer[:, :, 2] = 65

        mask_bool = mask_bin.astype(bool)
        overlay[mask_bool] = (
            0.55 * overlay[mask_bool].astype(np.float32) + 0.45 * red_layer[mask_bool].astype(np.float32)
        ).astype(np.uint8)

        contours, _ = cv2.findContours(mask_bin * 255, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        cv2.drawContours(overlay, contours, -1, (0, 255, 130), 2)
        x1, y1, x2, y2 = bbox
        cv2.rectangle(overlay, (x1, y1), (x2, y2), (255, 215, 0), 1)
        cv2.circle(overlay, peak_point, 3, (0, 220, 255), -1)

        return {
            "mask": mask_bin,
            "overlay_rgb": overlay,
            "bbox": bbox,
            "peak_point": peak_point,
            "tumor_pixels": tumor_pixels,
            "brain_pixels": brain_pixels,
            "tumor_ratio_pct": round(tumor_ratio_pct, 2),
            "method_used": method_used,
        }
