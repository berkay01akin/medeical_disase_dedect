"""
Çoklu Modalite & Derin Öğrenme Model Fabrikası (src/models.py)
--------------------------------------------------------------
3 farklı medikal görüntüleme modalitesi (Beyin MR, Göğüs Röntgeni, Akciğer BT) için
3 mimari ailesini ve 1 Otomatik Modalite Router (ResNet-18) modelini yönetir:
  - CNN:                ResNet-50 (resnet50)
  - Vision Transformer: ViT-B/16 (vit_base_patch16_224)
  - Swin Transformer:   Swin-T (swin_tiny_patch4_window7_224)
  - Modality Router:    ResNet-18 (brain_mri / chest_ct / chest_xray)
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, Tuple

import timm
import torch
import torch.nn as nn
from torchvision import transforms

PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = PROJECT_ROOT / "project_config.json"

IMAGENET_MEAN = (0.485, 0.456, 0.406)
IMAGENET_STD = (0.229, 0.224, 0.225)


def load_config() -> Dict[str, Any]:
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def get_modality_config(modality: str = "brain_mri") -> Dict[str, Any]:
    cfg = load_config()
    if modality not in cfg["modalities"]:
        raise ValueError(f"Bilinmeyen modalite: {modality}. Seçenekler: {list(cfg['modalities'].keys())}")
    return cfg["modalities"][modality]


def get_checkpoint_dir(modality: str = "brain_mri") -> Path:
    m_cfg = get_modality_config(modality)
    subdir = m_cfg.get("checkpoint_subdir", "")
    base = PROJECT_ROOT / "checkpoints"
    return (base / subdir) if subdir else base


def get_train_transforms(img_size: int = 224) -> transforms.Compose:
    """Eğitim sırasında veri artırma (Data Augmentation) dönüşümleri."""
    return transforms.Compose([
        transforms.Resize((img_size, img_size)),
        transforms.RandomHorizontalFlip(p=0.5),
        transforms.RandomRotation(degrees=15),
        transforms.RandomAffine(degrees=0, translate=(0.05, 0.05), scale=(0.95, 1.05)),
        transforms.ColorJitter(brightness=0.1, contrast=0.1),
        transforms.ToTensor(),
        transforms.Normalize(mean=IMAGENET_MEAN, std=IMAGENET_STD),
    ])


def get_eval_transforms(img_size: int = 224) -> transforms.Compose:
    """Doğrulama, test ve canlı tahmin dönüşümleri."""
    return transforms.Compose([
        transforms.Resize((img_size, img_size)),
        transforms.ToTensor(),
        transforms.Normalize(mean=IMAGENET_MEAN, std=IMAGENET_STD),
    ])


def create_classifier(model_key: str, num_classes: int = 4, pretrained: bool = True) -> nn.Module:
    """Seçilen anahtara ('cnn', 'vit', 'swin') göre timm sınıflandırıcı modeli oluşturur."""
    cfg = load_config()
    if model_key not in cfg["models"]:
        raise ValueError(f"Bilinmeyen model anahtarı: {model_key}. Seçenekler: {list(cfg['models'].keys())}")

    timm_name = cfg["models"][model_key]["timm_name"]
    return timm.create_model(timm_name, pretrained=pretrained, num_classes=num_classes)


def create_modality_router(num_modalities: int = 3, pretrained: bool = True) -> nn.Module:
    """MR / Röntgen / BT görüntüsünü otomatik ayıran hafif ResNet-18 yönlendirici modeli."""
    return timm.create_model("resnet18", pretrained=pretrained, num_classes=num_modalities)


def get_gradcam_target_layer(model: nn.Module, model_key: str) -> nn.Module:
    """Grad-CAM ısı haritası için mimariye uygun son öznitelik katmanını döndürür."""
    if model_key == "cnn":
        return model.layer4[-1]
    if model_key == "vit":
        return model.blocks[-1].norm1
    if model_key == "swin":
        return model.layers[-1].blocks[-1].norm1
    raise ValueError(f"Grad-CAM katmanı tanımlı değil: {model_key}")


def load_trained_model(
    model_key: str,
    checkpoint_dir: Path | str | None = None,
    device: torch.device | str | None = None,
    modality: str = "brain_mri",
) -> Tuple[nn.Module, Dict[str, Any]]:
    """Seçilen modaliteye ait eğitilmiş ağırlık dosyasını (.pth) yükler ve eval moduna alır."""
    m_cfg = get_modality_config(modality)
    ckpt_dir = Path(checkpoint_dir) if checkpoint_dir else get_checkpoint_dir(modality)
    ckpt_path = ckpt_dir / f"{model_key}_best.pth"
    if not ckpt_path.exists():
        raise FileNotFoundError(f"Model ağırlığı bulunamadı: {ckpt_path}")

    if device is None:
        device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    else:
        device = torch.device(device)

    payload = torch.load(ckpt_path, map_location=device, weights_only=False)
    num_classes = len(payload.get("classes", m_cfg["classes"]))

    model = create_classifier(model_key=model_key, num_classes=num_classes, pretrained=False)
    model.load_state_dict(payload["model_state_dict"])
    model.to(device)
    model.eval()
    return model, payload


def load_modality_router(
    checkpoint_dir: Path | str | None = None,
    device: torch.device | str | None = None,
) -> Tuple[nn.Module, Dict[str, Any]]:
    """Eğitilmiş Otomatik Modalite Router (checkpoints/router_best.pth) modelini yükler."""
    ckpt_dir = Path(checkpoint_dir) if checkpoint_dir else (PROJECT_ROOT / "checkpoints")
    ckpt_path = ckpt_dir / "router_best.pth"
    if not ckpt_path.exists():
        raise FileNotFoundError(f"Modalite Router ağırlığı bulunamadı: {ckpt_path}")

    if device is None:
        device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    else:
        device = torch.device(device)

    payload = torch.load(ckpt_path, map_location=device, weights_only=False)
    classes = payload["classes"]
    model = create_modality_router(num_modalities=len(classes), pretrained=False)
    model.load_state_dict(payload["model_state_dict"])
    model.to(device)
    model.eval()
    return model, payload
