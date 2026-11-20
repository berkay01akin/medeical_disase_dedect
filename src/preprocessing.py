"""
Klasik Görüntü İşleme Ön İşleme Modülü (src/preprocessing.py)
--------------------------------------------------------------
Medikal görüntülerde derin öğrenme öncesinde uygulanan temel görüntü işleme adımları:
1. Görüntü okuma (JPG/PNG ve DICOM .dcm desteği)
2. ROI (Region of Interest) kırpma:
   Gaussian Bulanıklaştırma -> Otsu Eşikleme -> Morfolojik Kapama/Açma -> En Büyük Kontur
3. Gürültü azaltma (Gaussian, Median, Bilateral filtreleri)
4. Kontrast iyileştirme (CLAHE - Contrast Limited Adaptive Histogram Equalization)
5. Histogram analizi ve boyutlandırma (224x224)
"""

from __future__ import annotations

import io
from pathlib import Path
from typing import Any, Dict, Tuple, Union

import cv2
import numpy as np
from PIL import Image


def load_medical_image(source: Union[str, Path, bytes, io.BytesIO, Image.Image, np.ndarray], filename: str = "") -> np.ndarray:
    """
    JPG, PNG, BMP veya DICOM (.dcm) formatındaki medikal görüntüyü RGB numpy dizisi (uint8) olarak yükler.
    """
    if isinstance(source, np.ndarray):
        img = source.copy()
        if img.ndim == 2:
            return cv2.cvtColor(img, cv2.COLOR_GRAY2RGB)
        if img.shape[2] == 4:
            return cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
        return img

    if isinstance(source, Image.Image):
        return np.array(source.convert("RGB"), dtype=np.uint8)

    is_dicom = filename.lower().endswith(".dcm")
    if isinstance(source, (str, Path)):
        path = Path(source)
        is_dicom = path.suffix.lower() == ".dcm"
        if is_dicom:
            return _read_dicom(path)
        pil_img = Image.open(path).convert("RGB")
        return np.array(pil_img, dtype=np.uint8)

    raw_bytes = source.getvalue() if isinstance(source, io.BytesIO) else source
    if is_dicom:
        return _read_dicom(io.BytesIO(raw_bytes))

    pil_img = Image.open(io.BytesIO(raw_bytes)).convert("RGB")
    return np.array(pil_img, dtype=np.uint8)


def _read_dicom(source: Union[Path, io.BytesIO]) -> np.ndarray:
    """DICOM (.dcm) dosyasını pydicom ile okuyup 8-bit RGB formatına normalize eder."""
    import pydicom

    ds = pydicom.dcmread(source)
    pixel_array = ds.pixel_array.astype(np.float32)

    slope = float(getattr(ds, "RescaleSlope", 1.0))
    intercept = float(getattr(ds, "RescaleIntercept", 0.0))
    pixel_array = pixel_array * slope + intercept

    p_min, p_max = np.percentile(pixel_array, (1, 99))
    if p_max > p_min:
        pixel_array = np.clip(pixel_array, p_min, p_max)
        pixel_array = (pixel_array - p_min) / (p_max - p_min) * 255.0
    else:
        pixel_array = np.zeros_like(pixel_array)

    gray = pixel_array.astype(np.uint8)
    return cv2.cvtColor(gray, cv2.COLOR_GRAY2RGB)


def extract_brain_roi(gray_img: np.ndarray, pad: int = 4) -> Tuple[np.ndarray, np.ndarray, np.ndarray, Tuple[int, int, int, int]]:
    """
    Otsu eşikleme + Morfolojik İşlemler (Kapama & Açma) + En Büyük Kontur ile
    görüntüdeki siyah arka plan boşluklarını temizler ve ilgi bölgesini (ROI) kırpar.
    """
    h, w = gray_img.shape[:2]
    blurred = cv2.GaussianBlur(gray_img, (5, 5), 0)

    # 1. Otsu otomatik eşikleme
    _, otsu_mask = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    # 2. Morfolojik işlemler: küçük gürültüleri sil (Açma) ve boşlukları doldur (Kapama)
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (7, 7))
    morph_mask = cv2.morphologyEx(otsu_mask, cv2.MORPH_CLOSE, kernel, iterations=2)
    morph_mask = cv2.morphologyEx(morph_mask, cv2.MORPH_OPEN, kernel, iterations=2)

    # 3. En büyük konturu bul
    contours, _ = cv2.findContours(morph_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return gray_img.copy(), otsu_mask, morph_mask, (0, 0, w, h)

    largest = max(contours, key=cv2.contourArea)
    area_ratio = cv2.contourArea(largest) / float(h * w)

    if area_ratio < 0.08:
        return gray_img.copy(), otsu_mask, morph_mask, (0, 0, w, h)

    x, y, bw, bh = cv2.boundingRect(largest)
    x1 = max(0, x - pad)
    y1 = max(0, y - pad)
    x2 = min(w, x + bw + pad)
    y2 = min(h, y + bh + pad)

    cropped = gray_img[y1:y2, x1:x2]
    return cropped, otsu_mask, morph_mask, (x1, y1, x2, y2)


def apply_noise_filter(gray_img: np.ndarray, filter_type: str = "bilateral", kernel_size: int = 5) -> np.ndarray:
    """
    Medikal görüntülerde gürültü azaltma filtreleri.
    - bilateral: Kenarları (tümör sınırlarını) koruyarak gürültü giderir.
    - median: Tuz-biber gürültüsüne karşı etkilidir.
    - gaussian: Yumuşak alçak geçiren filtre.
    - none: Filtre uygulamaz.
    """
    k = int(kernel_size)
    if k % 2 == 0:
        k += 1
    k = max(3, k)

    ft = filter_type.lower()
    if ft == "none":
        return gray_img.copy()
    if ft == "gaussian":
        return cv2.GaussianBlur(gray_img, (k, k), 0)
    if ft == "median":
        return cv2.medianBlur(gray_img, k)
    if ft == "bilateral":
        return cv2.bilateralFilter(gray_img, d=k, sigmaColor=50, sigmaSpace=50)
    raise ValueError(f"Desteklenmeyen filtre türü: {filter_type}")


def apply_clahe(gray_img: np.ndarray, clip_limit: float = 2.0, tile_grid_size: int = 8) -> np.ndarray:
    """
    CLAHE (Contrast Limited Adaptive Histogram Equalization):
    Yerel kontrastı artırarak doku ve tümör sınırlarını belirginleştirir.
    """
    grid = max(2, int(tile_grid_size))
    clahe = cv2.createCLAHE(clipLimit=float(clip_limit), tileGridSize=(grid, grid))
    return clahe.apply(gray_img)


def compute_histogram(gray_img: np.ndarray) -> np.ndarray:
    """256 seviyeli gri tonlama histogramı hesaplar."""
    return cv2.calcHist([gray_img], [0], None, [256], [0, 256]).flatten()


def preprocess_pipeline(
    image_input: Union[str, Path, bytes, io.BytesIO, Image.Image, np.ndarray],
    filename: str = "",
    use_roi: bool = True,
    filter_type: str = "bilateral",
    kernel_size: int = 5,
    use_clahe: bool = True,
    clahe_clip: float = 2.0,
    clahe_grid: int = 8,
    output_size: int = 224,
) -> Dict[str, Any]:
    """
    Uçtan uca klasik görüntü işleme hattı.
    Hem model girişi (224x224 RGB) hem de görselleştirme için tüm ara adımları döndürür.
    """
    original_rgb = load_medical_image(image_input, filename=filename)
    gray = cv2.cvtColor(original_rgb, cv2.COLOR_RGB2GRAY)

    # Adım 1: ROI Kırpma (Otsu + Morfoloji + Kontur)
    if use_roi:
        roi_gray, otsu_mask, morph_mask, roi_bbox = extract_brain_roi(gray)
    else:
        h, w = gray.shape[:2]
        roi_gray = gray.copy()
        otsu_mask = np.zeros_like(gray)
        morph_mask = np.zeros_like(gray)
        roi_bbox = (0, 0, w, h)

    hist_before = compute_histogram(roi_gray)

    # Adım 2: Gürültü Azaltma Filtresi
    filtered_gray = apply_noise_filter(roi_gray, filter_type=filter_type, kernel_size=kernel_size)

    # Adım 3: CLAHE Kontrast İyileştirme
    if use_clahe:
        enhanced_gray = apply_clahe(filtered_gray, clip_limit=clahe_clip, tile_grid_size=clahe_grid)
    else:
        enhanced_gray = filtered_gray.copy()

    hist_after = compute_histogram(enhanced_gray)

    # Adım 4: Boyutlandırma ve 3 kanallı RGB dönüşümü
    resized_gray = cv2.resize(enhanced_gray, (output_size, output_size), interpolation=cv2.INTER_AREA)
    final_rgb = cv2.cvtColor(resized_gray, cv2.COLOR_GRAY2RGB)

    x1, y1, x2, y2 = roi_bbox
    roi_rgb_resized = cv2.resize(original_rgb[y1:y2, x1:x2], (output_size, output_size), interpolation=cv2.INTER_AREA)

    return {
        "original_rgb": original_rgb,
        "gray": gray,
        "otsu_mask": otsu_mask,
        "morph_mask": morph_mask,
        "roi_bbox": roi_bbox,
        "roi_gray": roi_gray,
        "filtered_gray": filtered_gray,
        "enhanced_gray": enhanced_gray,
        "final_rgb": final_rgb,
        "roi_rgb_resized": roi_rgb_resized,
        "hist_before": hist_before,
        "hist_after": hist_after,
    }


def preprocess_for_training(image_path: Union[str, Path], output_size: int = 224) -> np.ndarray:
    """Eğitim verisi ön-hazırlığı için standart hızlı çağrı (ROI + Bilateral + CLAHE -> 224x224 RGB)."""
    result = preprocess_pipeline(
        image_input=image_path,
        use_roi=True,
        filter_type="bilateral",
        kernel_size=5,
        use_clahe=True,
        clahe_clip=2.0,
        clahe_grid=8,
        output_size=output_size,
    )
    return result["final_rgb"]
