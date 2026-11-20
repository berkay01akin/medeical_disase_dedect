"""
Çoklu Derin Öğrenme Model Fabrikası (src/models.py)
---------------------------------------------------
timm kütüphanesi ile 3 farklı mimari ailesini tek arayüzde birleştirir:
  1. CNN:                ResNet-50 (resnet50)
  2. Vision Transformer: ViT-B/16 (vit_base_patch16_224)
  3. Swin Transformer:   Swin-T (swin_tiny_patch4_window7_224)
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
) -> Tuple[nn.Module, Dict[str, Any]]:
    """Eğitilmiş ağırlık dosyasını (.pth) yükler ve eval moduna alır."""
    cfg = load_config()
    ckpt_dir = Path(checkpoint_dir) if checkpoint_dir else (PROJECT_ROOT / "checkpoints")
    ckpt_path = ckpt_dir / f"{model_key}_best.pth"
    if not ckpt_path.exists():
        raise FileNotFoundError(f"Model ağırlığı bulunamadı: {ckpt_path}")

    if device is None:
        device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    else:
        device = torch.device(device)

    payload = torch.load(ckpt_path, map_location=device, weights_only=False)
    num_classes = len(payload.get("classes", cfg["classes"]))

    model = create_classifier(model_key=model_key, num_classes=num_classes, pretrained=False)
    model.load_state_dict(payload["model_state_dict"])
    model.to(device)
    model.eval()
    return model, payload
