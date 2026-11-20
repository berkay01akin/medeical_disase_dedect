"""
Çoklu Model Eğitim, Değerlendirme ve Grafik Üretim Betiği (src/train.py)
------------------------------------------------------------------------
CNN (ResNet-50), Vision Transformer (ViT-B/16) ve Swin Transformer (Swin-T)
modellerini aynı ön-işlenmiş Beyin MR veri seti üzerinde eğitir, RTX 5070 Ti
üzerinde Mixed Precision (BF16 AMP) ile hızlandırır ve karşılaştırmalı
metrik + grafik çıktılarını (results/metrics.json, results/plots/*.png) üretir.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Tuple

import cv2
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns
import torch
import torch.nn as nn
from PIL import Image
from sklearn.metrics import (
    accuracy_score,
    classification_report,
    confusion_matrix,
    f1_score,
    precision_score,
    recall_score,
    roc_auc_score,
    roc_curve,
)
from sklearn.preprocessing import label_binarize
from torch.utils.data import DataLoader
from torchvision.datasets import ImageFolder
from tqdm import tqdm

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from src.explain_and_segment import TumorSegmenter, UnifiedGradCAM, mask_cam_to_brain, overlay_gradcam
from src.models import (
    create_classifier,
    get_eval_transforms,
    get_train_transforms,
    load_config,
    load_trained_model,
)
from src.preprocessing import preprocess_pipeline


def set_seed(seed: int = 42) -> None:
    np.random.seed(seed)
    torch.manual_seed(seed)
    if torch.cuda.is_available():
        torch.cuda.manual_seed_all(seed)


def train_one_epoch(
    model: nn.Module,
    loader: DataLoader,
    criterion: nn.Module,
    optimizer: torch.optim.Optimizer,
    device: torch.device,
    use_amp: bool,
) -> Tuple[float, float]:
    model.train()
    running_loss = 0.0
    correct = 0
    total = 0

    for images, labels in loader:
        images = images.to(device, non_blocking=True)
        labels = labels.to(device, non_blocking=True)

        optimizer.zero_grad(set_to_none=True)
        with torch.amp.autocast(device_type=device.type, dtype=torch.bfloat16, enabled=use_amp):
            outputs = model(images)
            loss = criterion(outputs, labels)

        loss.backward()
        torch.nn.utils.clip_grad_norm_(model.parameters(), max_norm=2.0)
        optimizer.step()

        running_loss += loss.item() * images.size(0)
        preds = outputs.argmax(dim=1)
        correct += (preds == labels).sum().item()
        total += labels.size(0)

    return running_loss / max(total, 1), correct / max(total, 1)


@torch.no_grad()
def evaluate_loader(
    model: nn.Module,
    loader: DataLoader,
    criterion: nn.Module,
    device: torch.device,
    use_amp: bool,
) -> Tuple[float, np.ndarray, np.ndarray, np.ndarray, float]:
    model.eval()
    running_loss = 0.0
    all_labels: List[np.ndarray] = []
    all_preds: List[np.ndarray] = []
    all_probs: List[np.ndarray] = []
    total = 0

    if device.type == "cuda":
        torch.cuda.synchronize()
    t0 = time.perf_counter()

    for images, labels in loader:
        images = images.to(device, non_blocking=True)
        labels = labels.to(device, non_blocking=True)

        with torch.amp.autocast(device_type=device.type, dtype=torch.bfloat16, enabled=use_amp):
            outputs = model(images)
            loss = criterion(outputs, labels)

        probs = torch.softmax(outputs.float(), dim=1)
        preds = probs.argmax(dim=1)

        running_loss += loss.item() * images.size(0)
        total += labels.size(0)
        all_labels.append(labels.cpu().numpy())
        all_preds.append(preds.cpu().numpy())
        all_probs.append(probs.cpu().numpy())

    if device.type == "cuda":
        torch.cuda.synchronize()
    elapsed_ms_per_img = ((time.perf_counter() - t0) / max(total, 1)) * 1000.0

    y_true = np.concatenate(all_labels)
    y_pred = np.concatenate(all_preds)
    y_prob = np.concatenate(all_probs)
    return running_loss / max(total, 1), y_true, y_pred, y_prob, elapsed_ms_per_img


def train_and_evaluate_model(
    model_key: str,
    cfg: Dict[str, Any],
    data_dir: Path,
    ckpt_dir: Path,
    device: torch.device,
    epochs_override: int | None = None,
) -> Tuple[Dict[str, Any], np.ndarray, np.ndarray, np.ndarray]:
    m_cfg = cfg["models"][model_key]
    epochs = epochs_override if epochs_override is not None else int(m_cfg["epochs"])
    batch_size = int(m_cfg["batch_size"])
    lr = float(m_cfg["lr"])
    use_amp = device.type == "cuda"

    train_ds = ImageFolder(data_dir / "train", transform=get_train_transforms(224))
    val_ds = ImageFolder(data_dir / "val", transform=get_eval_transforms(224))
    test_ds = ImageFolder(data_dir / "test", transform=get_eval_transforms(224))

    classes = train_ds.classes
    num_workers = 4 if device.type == "cuda" else 0
    train_loader = DataLoader(train_ds, batch_size=batch_size, shuffle=True, num_workers=num_workers, pin_memory=use_amp)
    val_loader = DataLoader(val_ds, batch_size=batch_size, shuffle=False, num_workers=num_workers, pin_memory=use_amp)
    test_loader = DataLoader(test_ds, batch_size=batch_size, shuffle=False, num_workers=num_workers, pin_memory=use_amp)

    print(f"\n{'='*72}")
    print(f"MODEL EĞİTİMİ BAŞLIYOR: {m_cfg['display_name']} ({m_cfg['timm_name']})")
    print(f"Cihaz: {device} | Epoch: {epochs} | Batch: {batch_size} | LR: {lr} | AMP BF16: {use_amp}")
    print(f"{'='*72}")

    model = create_classifier(model_key, num_classes=len(classes), pretrained=True).to(device)
    param_count_m = sum(p.numel() for p in model.parameters()) / 1e6

    criterion = nn.CrossEntropyLoss(label_smoothing=0.05)
    optimizer = torch.optim.AdamW(model.parameters(), lr=lr, weight_decay=1e-2)
    scheduler = torch.optim.lr_scheduler.CosineAnnealingLR(optimizer, T_max=epochs, eta_min=lr * 0.05)

    history = {"train_loss": [], "train_acc": [], "val_loss": [], "val_acc": [], "val_f1": []}
    best_val_f1 = -1.0
    best_epoch = 0
    ckpt_dir.mkdir(parents=True, exist_ok=True)
    ckpt_path = ckpt_dir / f"{model_key}_best.pth"

    t_start = time.perf_counter()
    for epoch in range(1, epochs + 1):
        ep_t0 = time.perf_counter()
        tr_loss, tr_acc = train_one_epoch(model, train_loader, criterion, optimizer, device, use_amp)
        val_loss, val_true, val_pred, _, _ = evaluate_loader(model, val_loader, criterion, device, use_amp)
        val_acc = accuracy_score(val_true, val_pred)
        val_f1 = f1_score(val_true, val_pred, average="macro")
        scheduler.step()

        history["train_loss"].append(round(float(tr_loss), 4))
        history["train_acc"].append(round(float(tr_acc), 4))
        history["val_loss"].append(round(float(val_loss), 4))
        history["val_acc"].append(round(float(val_acc), 4))
        history["val_f1"].append(round(float(val_f1), 4))

        marker = ""
        if val_f1 > best_val_f1:
            best_val_f1 = val_f1
            best_epoch = epoch
            torch.save(
                {
                    "model_key": model_key,
                    "timm_name": m_cfg["timm_name"],
                    "display_name": m_cfg["display_name"],
                    "classes": classes,
                    "best_epoch": best_epoch,
                    "best_val_f1": float(best_val_f1),
                    "model_state_dict": model.state_dict(),
                },
                ckpt_path,
            )
            marker = "  [* EN İYİ KAYDEDİLDİ]"

        ep_sec = time.perf_counter() - ep_t0
        print(
            f"  Epoch {epoch:02d}/{epochs:02d} ({ep_sec:4.1f}s) | "
            f"Train Loss: {tr_loss:.4f} Acc: {tr_acc*100:5.2f}% | "
            f"Val Loss: {val_loss:.4f} Acc: {val_acc*100:5.2f}% F1: {val_f1*100:5.2f}%{marker}"
        )

    train_time_sec = time.perf_counter() - t_start

    # En iyi checkpoint'i yükle ve TEST setinde değerlendir
    best_payload = torch.load(ckpt_path, map_location=device, weights_only=False)
    model.load_state_dict(best_payload["model_state_dict"])
    test_loss, y_true, y_pred, y_prob, ms_per_img = evaluate_loader(model, test_loader, criterion, device, use_amp)

    acc = accuracy_score(y_true, y_pred)
    prec = precision_score(y_true, y_pred, average="macro")
    rec = recall_score(y_true, y_pred, average="macro")
    f1 = f1_score(y_true, y_pred, average="macro")
    cm = confusion_matrix(y_true, y_pred)
    y_true_bin = label_binarize(y_true, classes=list(range(len(classes))))
    roc_auc = roc_auc_score(y_true_bin, y_prob, average="macro", multi_class="ovr")
    cls_report = classification_report(y_true, y_pred, target_names=classes, output_dict=True)

    print(
        f"--> TEST SONUCU [{m_cfg['short_name']}]: "
        f"Acc={acc*100:.2f}% | F1={f1*100:.2f}% | Precision={prec*100:.2f}% | "
        f"Recall={rec*100:.2f}% | ROC-AUC={roc_auc*100:.2f}% | {ms_per_img:.2f} ms/img"
    )

    metrics = {
        "model_key": model_key,
        "display_name": m_cfg["display_name"],
        "short_name": m_cfg["short_name"],
        "family": m_cfg["family"],
        "timm_name": m_cfg["timm_name"],
        "params_millions": round(float(param_count_m), 2),
        "train_time_sec": round(float(train_time_sec), 1),
        "best_epoch": int(best_epoch),
        "test_loss": round(float(test_loss), 4),
        "accuracy": round(float(acc), 4),
        "precision_macro": round(float(prec), 4),
        "recall_macro": round(float(rec), 4),
        "f1_macro": round(float(f1), 4),
        "roc_auc_macro": round(float(roc_auc), 4),
        "inference_ms_per_img": round(float(ms_per_img), 2),
        "confusion_matrix": cm.tolist(),
        "per_class": {
            cls: {
                "precision": round(float(cls_report[cls]["precision"]), 4),
                "recall": round(float(cls_report[cls]["recall"]), 4),
                "f1": round(float(cls_report[cls]["f1-score"]), 4),
                "support": int(cls_report[cls]["support"]),
            }
            for cls in classes
        },
        "history": history,
    }
    return metrics, y_true, y_pred, y_prob


def generate_comparison_plots(
    all_metrics: Dict[str, Dict[str, Any]],
    eval_outputs: Dict[str, Tuple[np.ndarray, np.ndarray, np.ndarray]],
    classes: List[str],
    plots_dir: Path,
) -> None:
    plots_dir.mkdir(parents=True, exist_ok=True)
    model_keys = list(all_metrics.keys())
    palette = {"cnn": "#38bdf8", "vit": "#a78bfa", "swin": "#34d399"}

    # 1. Eğitim Eğrileri (Loss & Val Accuracy)
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 4.8), dpi=180)
    fig.patch.set_facecolor("#0f172a")
    for ax in (ax1, ax2):
        ax.set_facecolor("#1e293b")
        ax.tick_params(colors="#cbd5e1")
        ax.grid(alpha=0.22, color="#64748b")

    for k in model_keys:
        m = all_metrics[k]
        ep = range(1, len(m["history"]["val_loss"]) + 1)
        color = palette.get(k, "#f472b6")
        ax1.plot(ep, m["history"]["train_loss"], linestyle="--", alpha=0.65, color=color, label=f"{m['short_name']} Train")
        ax1.plot(ep, m["history"]["val_loss"], linewidth=2.2, color=color, marker="o", markersize=4, label=f"{m['short_name']} Val")
        ax2.plot(ep, [v * 100 for v in m["history"]["val_acc"]], linewidth=2.4, color=color, marker="o", markersize=4, label=f"{m['short_name']} Val Acc")

    ax1.set_title("Eğitim & Doğrulama Kaybı (Loss)", color="white", fontweight="bold")
    ax1.set_xlabel("Epoch", color="#cbd5e1")
    ax1.set_ylabel("Loss", color="#cbd5e1")
    ax1.legend(facecolor="#0f172a", edgecolor="#334155", labelcolor="white", fontsize=8)

    ax2.set_title("Doğrulama Doğruluğu (Val Accuracy %)", color="white", fontweight="bold")
    ax2.set_xlabel("Epoch", color="#cbd5e1")
    ax2.set_ylabel("Accuracy (%)", color="#cbd5e1")
    ax2.legend(facecolor="#0f172a", edgecolor="#334155", labelcolor="white", fontsize=9)

    plt.tight_layout()
    fig.savefig(plots_dir / "training_curves.png", facecolor=fig.get_facecolor(), bbox_inches="tight")
    plt.close(fig)

    # 2. Yan Yana Confusion Matrices
    fig, axes = plt.subplots(1, len(model_keys), figsize=(5.2 * len(model_keys), 4.6), dpi=180)
    if len(model_keys) == 1:
        axes = [axes]
    fig.patch.set_facecolor("#0f172a")
    short_cls = ["Glioma", "Menenjiyom", "Tümör Yok", "Hipofiz"]

    for ax, k in zip(axes, model_keys):
        m = all_metrics[k]
        cm = np.array(m["confusion_matrix"])
        sns.heatmap(
            cm,
            annot=True,
            fmt="d",
            cmap="Blues",
            cbar=False,
            xticklabels=short_cls,
            yticklabels=short_cls,
            ax=ax,
            annot_kws={"size": 11, "weight": "bold"},
        )
        ax.set_title(f"{m['short_name']} (Acc: {m['accuracy']*100:.2f}%)", color="white", fontsize=11, fontweight="bold", pad=10)
        ax.set_xlabel("Tahmin", color="#cbd5e1")
        ax.set_ylabel("Gerçek", color="#cbd5e1")
        ax.tick_params(colors="#e2e8f0", labelsize=9)

    fig.suptitle("Test Seti Karışıklık Matrisleri (Confusion Matrix — 1.600 MR)", color="#38bdf8", fontsize=13, fontweight="bold", y=1.02)
    plt.tight_layout()
    fig.savefig(plots_dir / "confusion_matrices.png", facecolor=fig.get_facecolor(), bbox_inches="tight")
    plt.close(fig)

    # 3. Model Karşılaştırma Bar Grafiği
    metric_names = ["Accuracy", "Precision", "Recall", "Macro F1", "ROC-AUC"]
    metric_keys = ["accuracy", "precision_macro", "recall_macro", "f1_macro", "roc_auc_macro"]
    x = np.arange(len(metric_names))
    width = 0.25

    fig, ax = plt.subplots(figsize=(10.5, 5), dpi=180)
    fig.patch.set_facecolor("#0f172a")
    ax.set_facecolor("#1e293b")

    for i, k in enumerate(model_keys):
        m = all_metrics[k]
        vals = [m[mk] * 100 for mk in metric_keys]
        offset = (i - (len(model_keys) - 1) / 2) * width
        bars = ax.bar(x + offset, vals, width, label=m["display_name"], color=palette.get(k, "#f472b6"), edgecolor="#0f172a")
        for bar in bars:
            h = bar.get_height()
            ax.annotate(f"{h:.1f}%", xy=(bar.get_x() + bar.get_width() / 2, h), xytext=(0, 3), textcoords="offset points", ha="center", va="bottom", color="white", fontsize=8, fontweight="bold")

    ax.set_ylim(max(70, min(all_metrics[k]["accuracy"] * 100 for k in model_keys) - 6), 101.5)
    ax.set_xticks(x)
    ax.set_xticklabels(metric_names, color="white", fontsize=10, fontweight="bold")
    ax.set_ylabel("Skor (%)", color="#cbd5e1")
    ax.set_title("CNN vs ViT vs Swin Transformer — Test Seti Performans Karşılaştırması", color="#38bdf8", fontsize=13, fontweight="bold", pad=12)
    ax.tick_params(colors="#cbd5e1")
    ax.grid(axis="y", alpha=0.22, color="#64748b")
    ax.legend(facecolor="#0f172a", edgecolor="#334155", labelcolor="white", loc="lower right")
    plt.tight_layout()
    fig.savefig(plots_dir / "model_comparison.png", facecolor=fig.get_facecolor(), bbox_inches="tight")
    plt.close(fig)

    # 4. Macro-Average ROC Eğrileri
    fig, ax = plt.subplots(figsize=(7.5, 5.5), dpi=180)
    fig.patch.set_facecolor("#0f172a")
    ax.set_facecolor("#1e293b")
    for k in model_keys:
        if k not in eval_outputs:
            continue
        y_true, _, y_prob = eval_outputs[k]
        y_bin = label_binarize(y_true, classes=list(range(len(classes))))
        fpr, tpr, _ = roc_curve(y_bin.ravel(), y_prob.ravel())
        m = all_metrics[k]
        ax.plot(fpr, tpr, linewidth=2.3, color=palette.get(k, "#f472b6"), label=f"{m['short_name']} (AUC = {m['roc_auc_macro']*100:.2f}%)")

    ax.plot([0, 1], [0, 1], linestyle="--", color="#64748b", linewidth=1.2)
    ax.set_xlabel("False Positive Rate", color="#cbd5e1")
    ax.set_ylabel("True Positive Rate", color="#cbd5e1")
    ax.set_title("Çok Sınıflı ROC Eğrileri (Micro/Macro Karşılaştırma)", color="white", fontweight="bold")
    ax.tick_params(colors="#cbd5e1")
    ax.grid(alpha=0.2, color="#64748b")
    ax.legend(facecolor="#0f172a", edgecolor="#334155", labelcolor="white", loc="lower right")
    plt.tight_layout()
    fig.savefig(plots_dir / "roc_curves.png", facecolor=fig.get_facecolor(), bbox_inches="tight")
    plt.close(fig)


def generate_gradcam_sam_showcase(cfg: Dict[str, Any], best_model_key: str, device: torch.device, plots_dir: Path) -> None:
    """Sunum ve rapor için 3 farklı tümör tipinde Grad-CAM + SAM Segmentasyon vitrin görseli üretir."""
    sample_dir = PROJECT_ROOT / "sample_images"
    sample_files = [
        ("glioma", sample_dir / "glioma_01.jpg"),
        ("meningioma", sample_dir / "meningioma_01.jpg"),
        ("pituitary", sample_dir / "pituitary_01.jpg"),
    ]
    valid_samples = [(cls, p) for cls, p in sample_files if p.exists()]
    if not valid_samples:
        return

    model, payload = load_trained_model(best_model_key, device=device)
    classes = payload["classes"]
    transform = get_eval_transforms(224)
    # CNN (ResNet-50) uzamsal Grad-CAM yerel tümör lokalizasyonunda çok daha keskindir;
    # sınıflandırma tahmini en iyi modelle (Swin-T), ısı haritası CNN Grad-CAM ile üretilir.
    cam_model_key = "cnn" if (PROJECT_ROOT / "checkpoints" / "cnn_best.pth").exists() else best_model_key
    cam_model, _ = load_trained_model(cam_model_key, device=device)
    gradcam = UnifiedGradCAM(cam_model, cam_model_key)
    segmenter = TumorSegmenter()

    fig, axes = plt.subplots(len(valid_samples), 4, figsize=(13, 3.3 * len(valid_samples)), dpi=180)
    fig.patch.set_facecolor("#0f172a")
    col_titles = [
        "Orijinal MR",
        "CLAHE + ROI Ön İşleme",
        f"Grad-CAM ({cfg['models'][cam_model_key]['short_name']})",
        "SAM Tümör Segmentasyonu",
    ]

    for row_idx, (true_cls, img_path) in enumerate(valid_samples):
        steps = preprocess_pipeline(img_path)
        inp_tensor = transform(Image.fromarray(steps["final_rgb"])).unsqueeze(0).to(device)
        with torch.no_grad():
            probs = torch.softmax(model(inp_tensor), dim=1)[0].cpu().numpy()
            pred_idx = int(np.argmax(probs))

        cam_2d, _, _ = gradcam.generate(inp_tensor, target_class=pred_idx)
        clean_cam = mask_cam_to_brain(cam_2d, steps["roi_rgb_resized"], erode_px=12)
        cam_overlay = overlay_gradcam(steps["final_rgb"], clean_cam, alpha=0.45)
        seg = segmenter.segment(steps["roi_rgb_resized"], clean_cam, cam_threshold=0.52)

        pred_cls = classes[pred_idx]
        conf = probs[pred_idx] * 100.0
        row_imgs = [steps["original_rgb"], steps["final_rgb"], cam_overlay, seg["overlay_rgb"]]
        row_sub = [
            f"Gerçek: {cfg['class_labels_tr'][true_cls]}",
            "Bilateral + CLAHE (224x224)",
            f"Tahmin: {cfg['class_labels_tr'][pred_cls]} (%{conf:.1f})",
            f"SAM Alan Oranı: %{seg['tumor_ratio_pct']:.1f}",
        ]

        for col_idx in range(4):
            ax = axes[row_idx, col_idx] if len(valid_samples) > 1 else axes[col_idx]
            ax.imshow(row_imgs[col_idx])
            if row_idx == 0:
                ax.set_title(col_titles[col_idx], color="#38bdf8", fontsize=11, fontweight="bold", pad=8)
            ax.set_xlabel(row_sub[col_idx], color="#e2e8f0", fontsize=9)
            ax.set_xticks([])
            ax.set_yticks([])

    gradcam.remove_hooks()
    fig.suptitle("Açıklanabilirlik (Grad-CAM) ve SAM Tabanlı Otomatik Tümör Segmentasyonu", color="white", fontsize=13, fontweight="bold", y=0.99)
    plt.tight_layout(rect=[0, 0, 1, 0.96])
    fig.savefig(plots_dir / "gradcam_sam_showcase.png", facecolor=fig.get_facecolor(), bbox_inches="tight")
    plt.close(fig)


def main() -> None:
    parser = argparse.ArgumentParser(description="CNN, ViT ve Swin Transformer modellerini eğitir ve kıyaslar.")
    parser.add_argument("--models", nargs="+", default=["cnn", "vit", "swin"], choices=["cnn", "vit", "swin"])
    parser.add_argument("--epochs", type=int, default=None, help="Tüm modeller için epoch sayısını geçersiz kılar")
    parser.add_argument("--seed", type=int, default=42)
    args = parser.parse_args()

    set_seed(args.seed)
    cfg = load_config()
    device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    print(f"Kullanılan Cihaz: {device}" + (f" ({torch.cuda.get_device_name(0)})" if device.type == "cuda" else ""))

    data_dir = PROJECT_ROOT / "data" / "processed"
    ckpt_dir = PROJECT_ROOT / "checkpoints"
    results_dir = PROJECT_ROOT / "results"
    plots_dir = results_dir / "plots"
    metrics_path = results_dir / "metrics.json"

    all_metrics: Dict[str, Dict[str, Any]] = {}
    if metrics_path.exists():
        with open(metrics_path, "r", encoding="utf-8") as f:
            existing = json.load(f)
            all_metrics = existing.get("models", {})

    eval_outputs: Dict[str, Tuple[np.ndarray, np.ndarray, np.ndarray]] = {}

    for mk in args.models:
        m_metrics, y_true, y_pred, y_prob = train_and_evaluate_model(
            model_key=mk,
            cfg=cfg,
            data_dir=data_dir,
            ckpt_dir=ckpt_dir,
            device=device,
            epochs_override=args.epochs,
        )
        all_metrics[mk] = m_metrics
        eval_outputs[mk] = (y_true, y_pred, y_prob)

    best_key = max(all_metrics.keys(), key=lambda k: all_metrics[k]["f1_macro"])
    payload = {
        "device": torch.cuda.get_device_name(0) if device.type == "cuda" else "CPU",
        "classes": cfg["classes"],
        "best_model_key": best_key,
        "models": all_metrics,
    }
    results_dir.mkdir(parents=True, exist_ok=True)
    with open(metrics_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    print("\nKarşılaştırma grafikleri üretiliyor...")
    generate_comparison_plots(all_metrics, eval_outputs, cfg["classes"], plots_dir)
    print("Grad-CAM + SAM vitrin görseli üretiliyor...")
    generate_gradcam_sam_showcase(cfg, best_key, device, plots_dir)
    print(f"\nTAMAMLANDI! En iyi model: {all_metrics[best_key]['display_name']} (F1={all_metrics[best_key]['f1_macro']*100:.2f}%)")
    print(f"Sonuçlar: {metrics_path}")


if __name__ == "__main__":
    main()
