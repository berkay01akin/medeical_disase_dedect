"""
Çoklu Modalite ve Otomatik Modalite Router Eğitim Betiği (src/train_multimodal.py)
----------------------------------------------------------------------------------
1. router     -> ResNet-18 ile Beyin MR / Göğüs Röntgeni / Akciğer BT otomatik algılama
2. chest_ct   -> Akciğer Kanseri BT (4 sınıf) için CNN (ResNet-50), ViT-B/16, Swin-T eğitimi
3. chest_xray -> Göğüs Röntgeni COVID-19/Pnömoni/TB/Normal için CNN, ViT-B/16, Swin-T eğitimi
4. Görseller  -> Her modalite için Confusion Matrix, Model Comparison, Training Curves, ROC-AUC
                 + 3 Modaliteli Genel Özet ve Grad-CAM + MobileSAM Vitrin Grafiği üretir.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Tuple

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

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from src.explain_and_segment import TumorSegmenter, UnifiedGradCAM, mask_cam_to_brain, overlay_gradcam
from src.models import (
    create_classifier,
    create_modality_router,
    get_checkpoint_dir,
    get_eval_transforms,
    get_modality_config,
    get_train_transforms,
    load_config,
    load_trained_model,
)
from src.preprocessing import preprocess_pipeline
from src.train import evaluate_loader, set_seed, train_one_epoch


def train_router(device: torch.device, epochs: int = 5) -> Dict[str, Any]:
    data_dir = PROJECT_ROOT / "data" / "router"
    ckpt_path = PROJECT_ROOT / "checkpoints" / "router_best.pth"
    use_amp = device.type == "cuda"

    train_ds = ImageFolder(data_dir / "train", transform=get_train_transforms(224))
    val_ds = ImageFolder(data_dir / "val", transform=get_eval_transforms(224))
    test_ds = ImageFolder(data_dir / "test", transform=get_eval_transforms(224))
    classes = train_ds.classes

    num_workers = 4 if use_amp else 0
    train_loader = DataLoader(train_ds, batch_size=32, shuffle=True, num_workers=num_workers, pin_memory=use_amp)
    val_loader = DataLoader(val_ds, batch_size=32, shuffle=False, num_workers=num_workers, pin_memory=use_amp)
    test_loader = DataLoader(test_ds, batch_size=32, shuffle=False, num_workers=num_workers, pin_memory=use_amp)

    print(f"\n{'='*72}")
    print(f"OTOMATİK MODALİTE ROUTER EĞİTİMİ (ResNet-18) — Sınıflar: {classes}")
    print(f"{'='*72}")

    model = create_modality_router(num_modalities=len(classes), pretrained=True).to(device)
    criterion = nn.CrossEntropyLoss(label_smoothing=0.02)
    optimizer = torch.optim.AdamW(model.parameters(), lr=3e-4, weight_decay=1e-2)
    scheduler = torch.optim.lr_scheduler.CosineAnnealingLR(optimizer, T_max=epochs)

    best_f1 = -1.0
    t0 = time.perf_counter()
    for ep in range(1, epochs + 1):
        tr_loss, tr_acc = train_one_epoch(model, train_loader, criterion, optimizer, device, use_amp)
        val_loss, v_true, v_pred, _, _ = evaluate_loader(model, val_loader, criterion, device, use_amp)
        val_acc = accuracy_score(v_true, v_pred)
        val_f1 = f1_score(v_true, v_pred, average="macro")
        scheduler.step()
        mark = ""
        if val_f1 > best_f1:
            best_f1 = val_f1
            ckpt_path.parent.mkdir(parents=True, exist_ok=True)
            torch.save({"classes": classes, "model_state_dict": model.state_dict(), "val_f1": float(best_f1)}, ckpt_path)
            mark = " [* EN İYİ]"
        print(f"  Router Ep {ep:02d}/{epochs:02d} | TrAcc: {tr_acc*100:5.2f}% | ValAcc: {val_acc*100:5.2f}% F1: {val_f1*100:5.2f}%{mark}")

    payload = torch.load(ckpt_path, map_location=device, weights_only=False)
    model.load_state_dict(payload["model_state_dict"])
    _, y_true, y_pred, _, ms_img = evaluate_loader(model, test_loader, criterion, device, use_amp)
    acc = accuracy_score(y_true, y_pred)
    f1 = f1_score(y_true, y_pred, average="macro")
    cm = confusion_matrix(y_true, y_pred)

    metrics = {
        "model": "ResNet-18 (Otomatik Modalite Router)",
        "classes": classes,
        "test_samples": int(len(y_true)),
        "accuracy": round(float(acc), 4),
        "f1_macro": round(float(f1), 4),
        "inference_ms_per_img": round(float(ms_img), 2),
        "train_time_sec": round(float(time.perf_counter() - t0), 1),
        "confusion_matrix": cm.tolist(),
    }
    out_json = PROJECT_ROOT / "results" / "router_metrics.json"
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(metrics, f, ensure_ascii=False, indent=2)
    print(f"--> ROUTER TEST SONUCU: Acc={acc*100:.2f}% | F1={f1*100:.2f}% | {ms_img:.2f} ms/img")
    return metrics


def plot_modality_results(
    modality: str,
    m_cfg: Dict[str, Any],
    all_metrics: Dict[str, Dict[str, Any]],
    eval_outputs: Dict[str, Tuple[np.ndarray, np.ndarray, np.ndarray]],
    classes: List[str],
    plots_dir: Path,
) -> None:
    plots_dir.mkdir(parents=True, exist_ok=True)
    model_keys = list(all_metrics.keys())
    palette = {"cnn": "#38bdf8", "vit": "#a78bfa", "swin": "#34d399"}
    short_labels = [m_cfg["class_labels_tr"][c].split(" (")[0] for c in classes]

    # 1. Model Karşılaştırma Bar Grafiği
    metric_names = ["Accuracy", "Precision", "Recall", "Macro F1", "ROC-AUC"]
    metric_keys = ["accuracy", "precision_macro", "recall_macro", "f1_macro", "roc_auc_macro"]
    x = np.arange(len(metric_names))
    width = 0.25

    fig, ax = plt.subplots(figsize=(10.5, 4.8), dpi=180)
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
    min_score = min(all_metrics[k]["f1_macro"] * 100 for k in model_keys)
    ax.set_ylim(max(50, min_score - 10), 102.0)
    ax.set_xticks(x)
    ax.set_xticklabels(metric_names, color="white", fontsize=10, fontweight="bold")
    ax.set_ylabel("Skor (%)", color="#cbd5e1")
    ax.set_title(f"{m_cfg['display_name']} — CNN vs ViT vs Swin Performans Karşılaştırması", color="#38bdf8", fontsize=12.5, fontweight="bold", pad=12)
    ax.tick_params(colors="#cbd5e1")
    ax.grid(axis="y", alpha=0.22, color="#64748b")
    ax.legend(facecolor="#0f172a", edgecolor="#334155", labelcolor="white", loc="lower right")
    plt.tight_layout()
    fig.savefig(plots_dir / "model_comparison.png", facecolor=fig.get_facecolor(), bbox_inches="tight")
    plt.close(fig)

    # 2. Confusion Matrices
    fig, axes = plt.subplots(1, len(model_keys), figsize=(5.3 * len(model_keys), 4.6), dpi=180)
    if len(model_keys) == 1:
        axes = [axes]
    fig.patch.set_facecolor("#0f172a")
    for ax, k in zip(axes, model_keys):
        m = all_metrics[k]
        cm = np.array(m["confusion_matrix"])
        sns.heatmap(cm, annot=True, fmt="d", cmap="Blues", cbar=False, xticklabels=short_labels, yticklabels=short_labels, ax=ax, annot_kws={"size": 10.5, "weight": "bold"})
        ax.set_title(f"{m['short_name']} (Acc: {m['accuracy']*100:.2f}%)", color="white", fontsize=11, fontweight="bold", pad=10)
        ax.set_xlabel("Tahmin", color="#cbd5e1")
        ax.set_ylabel("Gerçek", color="#cbd5e1")
        ax.tick_params(colors="#e2e8f0", labelsize=8.5, rotation=15)
    fig.suptitle(f"{m_cfg['short_name']} — Test Seti Karışıklık Matrisleri", color="#38bdf8", fontsize=13, fontweight="bold", y=1.02)
    plt.tight_layout()
    fig.savefig(plots_dir / "confusion_matrices.png", facecolor=fig.get_facecolor(), bbox_inches="tight")
    plt.close(fig)

    # 3. Training Curves
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12.5, 4.5), dpi=180)
    fig.patch.set_facecolor("#0f172a")
    for ax in (ax1, ax2):
        ax.set_facecolor("#1e293b")
        ax.tick_params(colors="#cbd5e1")
        ax.grid(alpha=0.22, color="#64748b")
    for k in model_keys:
        m = all_metrics[k]
        ep = range(1, len(m["history"]["val_loss"]) + 1)
        col = palette.get(k, "#f472b6")
        ax1.plot(ep, m["history"]["train_loss"], "--", alpha=0.6, color=col, label=f"{m['short_name']} Train")
        ax1.plot(ep, m["history"]["val_loss"], linewidth=2.2, marker="o", markersize=3.5, color=col, label=f"{m['short_name']} Val")
        ax2.plot(ep, [v * 100 for v in m["history"]["val_acc"]], linewidth=2.2, marker="o", markersize=3.5, color=col, label=f"{m['short_name']} Val Acc")
    ax1.set_title(f"{m_cfg['short_name']} — Kayıp Eğrileri (Loss)", color="white", fontweight="bold")
    ax1.set_xlabel("Epoch", color="#cbd5e1")
    ax1.set_ylabel("Loss", color="#cbd5e1")
    ax1.legend(facecolor="#0f172a", edgecolor="#334155", labelcolor="white", fontsize=8)
    ax2.set_title(f"{m_cfg['short_name']} — Doğrulama Doğruluğu (%)", color="white", fontweight="bold")
    ax2.set_xlabel("Epoch", color="#cbd5e1")
    ax2.set_ylabel("Accuracy (%)", color="#cbd5e1")
    ax2.legend(facecolor="#0f172a", edgecolor="#334155", labelcolor="white", fontsize=9)
    plt.tight_layout()
    fig.savefig(plots_dir / "training_curves.png", facecolor=fig.get_facecolor(), bbox_inches="tight")
    plt.close(fig)

    # 4. ROC Curves
    fig, ax = plt.subplots(figsize=(7.2, 5.2), dpi=180)
    fig.patch.set_facecolor("#0f172a")
    ax.set_facecolor("#1e293b")
    for k in model_keys:
        if k not in eval_outputs:
            continue
        y_true, _, y_prob = eval_outputs[k]
        y_bin = label_binarize(y_true, classes=list(range(len(classes))))
        fpr, tpr, _ = roc_curve(y_bin.ravel(), y_prob.ravel())
        m = all_metrics[k]
        ax.plot(fpr, tpr, linewidth=2.2, color=palette.get(k, "#f472b6"), label=f"{m['short_name']} (AUC = {m['roc_auc_macro']*100:.2f}%)")
    ax.plot([0, 1], [0, 1], "--", color="#64748b", linewidth=1.2)
    ax.set_xlabel("False Positive Rate", color="#cbd5e1")
    ax.set_ylabel("True Positive Rate", color="#cbd5e1")
    ax.set_title(f"{m_cfg['short_name']} — ROC Eğrileri", color="white", fontweight="bold")
    ax.tick_params(colors="#cbd5e1")
    ax.grid(alpha=0.2, color="#64748b")
    ax.legend(facecolor="#0f172a", edgecolor="#334155", labelcolor="white", loc="lower right")
    plt.tight_layout()
    fig.savefig(plots_dir / "roc_curves.png", facecolor=fig.get_facecolor(), bbox_inches="tight")
    plt.close(fig)


def train_modality(modality: str, device: torch.device, epochs: int) -> Dict[str, Any]:
    cfg = load_config()
    m_cfg = get_modality_config(modality)
    data_dir = PROJECT_ROOT / "data" / modality / "processed"
    ckpt_dir = get_checkpoint_dir(modality)
    plots_dir = PROJECT_ROOT / "results" / "plots" / m_cfg["plots_subdir"]
    metrics_path = PROJECT_ROOT / "results" / m_cfg["metrics_file"]
    use_amp = device.type == "cuda"

    train_ds = ImageFolder(data_dir / "train", transform=get_train_transforms(224))
    val_ds = ImageFolder(data_dir / "val", transform=get_eval_transforms(224))
    test_ds = ImageFolder(data_dir / "test", transform=get_eval_transforms(224))
    classes = train_ds.classes

    num_workers = 4 if use_amp else 0
    all_metrics: Dict[str, Dict[str, Any]] = {}
    eval_outputs: Dict[str, Tuple[np.ndarray, np.ndarray, np.ndarray]] = {}

    for mk in ["cnn", "vit", "swin"]:
        model_meta = cfg["models"][mk]
        batch_size = int(model_meta["batch_size"])
        lr = float(model_meta["lr"])

        train_loader = DataLoader(train_ds, batch_size=batch_size, shuffle=True, num_workers=num_workers, pin_memory=use_amp)
        val_loader = DataLoader(val_ds, batch_size=batch_size, shuffle=False, num_workers=num_workers, pin_memory=use_amp)
        test_loader = DataLoader(test_ds, batch_size=batch_size, shuffle=False, num_workers=num_workers, pin_memory=use_amp)

        print(f"\n{'='*72}")
        print(f"[{m_cfg['short_name'].upper()}] {model_meta['display_name']} EĞİTİMİ ({epochs} Epoch)")
        print(f"{'='*72}")

        model = create_classifier(mk, num_classes=len(classes), pretrained=True).to(device)
        param_m = sum(p.numel() for p in model.parameters()) / 1e6
        criterion = nn.CrossEntropyLoss(label_smoothing=0.05)
        optimizer = torch.optim.AdamW(model.parameters(), lr=lr, weight_decay=1e-2)
        scheduler = torch.optim.lr_scheduler.CosineAnnealingLR(optimizer, T_max=epochs, eta_min=lr * 0.05)

        history = {"train_loss": [], "train_acc": [], "val_loss": [], "val_acc": [], "val_f1": []}
        best_val_f1 = -1.0
        best_epoch = 0
        ckpt_dir.mkdir(parents=True, exist_ok=True)
        ckpt_path = ckpt_dir / f"{mk}_best.pth"

        t0 = time.perf_counter()
        for ep in range(1, epochs + 1):
            ep_t0 = time.perf_counter()
            tr_loss, tr_acc = train_one_epoch(model, train_loader, criterion, optimizer, device, use_amp)
            val_loss, v_true, v_pred, _, _ = evaluate_loader(model, val_loader, criterion, device, use_amp)
            val_acc = accuracy_score(v_true, v_pred)
            val_f1 = f1_score(v_true, v_pred, average="macro")
            scheduler.step()

            history["train_loss"].append(round(float(tr_loss), 4))
            history["train_acc"].append(round(float(tr_acc), 4))
            history["val_loss"].append(round(float(val_loss), 4))
            history["val_acc"].append(round(float(val_acc), 4))
            history["val_f1"].append(round(float(val_f1), 4))

            mark = ""
            if val_f1 > best_val_f1:
                best_val_f1 = val_f1
                best_epoch = ep
                torch.save(
                    {
                        "modality": modality,
                        "model_key": mk,
                        "timm_name": model_meta["timm_name"],
                        "display_name": model_meta["display_name"],
                        "classes": classes,
                        "best_epoch": best_epoch,
                        "best_val_f1": float(best_val_f1),
                        "model_state_dict": model.state_dict(),
                    },
                    ckpt_path,
                )
                mark = " [* EN İYİ]"
            print(f"  Ep {ep:02d}/{epochs:02d} ({time.perf_counter()-ep_t0:4.1f}s) | TrAcc: {tr_acc*100:5.2f}% | ValAcc: {val_acc*100:5.2f}% F1: {val_f1*100:5.2f}%{mark}")

        train_sec = time.perf_counter() - t0
        best_payload = torch.load(ckpt_path, map_location=device, weights_only=False)
        model.load_state_dict(best_payload["model_state_dict"])
        test_loss, y_true, y_pred, y_prob, ms_img = evaluate_loader(model, test_loader, criterion, device, use_amp)

        acc = accuracy_score(y_true, y_pred)
        prec = precision_score(y_true, y_pred, average="macro", zero_division=0)
        rec = recall_score(y_true, y_pred, average="macro", zero_division=0)
        f1 = f1_score(y_true, y_pred, average="macro", zero_division=0)
        cm = confusion_matrix(y_true, y_pred)
        y_bin = label_binarize(y_true, classes=list(range(len(classes))))
        auc = roc_auc_score(y_bin, y_prob, average="macro", multi_class="ovr")
        cls_rep = classification_report(y_true, y_pred, target_names=classes, output_dict=True, zero_division=0)

        print(f"--> [{m_cfg['short_name']} | {model_meta['short_name']}] TEST: Acc={acc*100:.2f}% | F1={f1*100:.2f}% | AUC={auc*100:.2f}% | {ms_img:.2f} ms/img")

        all_metrics[mk] = {
            "model_key": mk,
            "display_name": model_meta["display_name"],
            "short_name": model_meta["short_name"],
            "family": model_meta["family"],
            "timm_name": model_meta["timm_name"],
            "params_millions": round(float(param_m), 2),
            "train_time_sec": round(float(train_sec), 1),
            "best_epoch": int(best_epoch),
            "test_loss": round(float(test_loss), 4),
            "accuracy": round(float(acc), 4),
            "precision_macro": round(float(prec), 4),
            "recall_macro": round(float(rec), 4),
            "f1_macro": round(float(f1), 4),
            "roc_auc_macro": round(float(auc), 4),
            "inference_ms_per_img": round(float(ms_img), 2),
            "test_samples": int(len(y_true)),
            "confusion_matrix": cm.tolist(),
            "per_class": {
                c: {
                    "precision": round(float(cls_rep[c]["precision"]), 4),
                    "recall": round(float(cls_rep[c]["recall"]), 4),
                    "f1": round(float(cls_rep[c]["f1-score"]), 4),
                    "support": int(cls_rep[c]["support"]),
                }
                for c in classes
            },
            "history": history,
        }
        eval_outputs[mk] = (y_true, y_pred, y_prob)

    best_key = max(all_metrics.keys(), key=lambda k: all_metrics[k]["f1_macro"])
    result_payload = {
        "modality": modality,
        "display_name": m_cfg["display_name"],
        "device": torch.cuda.get_device_name(0) if device.type == "cuda" else "CPU",
        "classes": classes,
        "best_model_key": best_key,
        "models": all_metrics,
    }
    with open(metrics_path, "w", encoding="utf-8") as f:
        json.dump(result_payload, f, ensure_ascii=False, indent=2)

    plot_modality_results(modality, m_cfg, all_metrics, eval_outputs, classes, plots_dir)
    return result_payload


def generate_multimodal_summary_plots(device: torch.device) -> None:
    cfg = load_config()
    plots_dir = PROJECT_ROOT / "results" / "plots"

    # 1. 3 Modalite Karşılaştırma Genel Grafiği
    mod_keys = ["brain_mri", "chest_xray", "chest_ct"]
    mod_names = []
    cnn_f1, vit_f1, swin_f1 = [], [], []
    for mod in mod_keys:
        m_cfg = cfg["modalities"][mod]
        m_path = PROJECT_ROOT / "results" / m_cfg["metrics_file"]
        if not m_path.exists():
            continue
        data = json.load(open(m_path, encoding="utf-8"))
        mod_names.append(m_cfg["short_name"])
        cnn_f1.append(data["models"]["cnn"]["f1_macro"] * 100)
        vit_f1.append(data["models"]["vit"]["f1_macro"] * 100)
        swin_f1.append(data["models"]["swin"]["f1_macro"] * 100)

    if len(mod_names) == 3:
        x = np.arange(3)
        w = 0.24
        fig, ax = plt.subplots(figsize=(10.5, 4.8), dpi=180)
        fig.patch.set_facecolor("#0f172a")
        ax.set_facecolor("#1e293b")
        b1 = ax.bar(x - w, cnn_f1, w, label="CNN (ResNet-50)", color="#38bdf8", edgecolor="#0f172a")
        b2 = ax.bar(x, vit_f1, w, label="Vision Transformer (ViT-B/16)", color="#a78bfa", edgecolor="#0f172a")
        b3 = ax.bar(x + w, swin_f1, w, label="Swin Transformer (Swin-T)", color="#34d399", edgecolor="#0f172a")
        for bars in (b1, b2, b3):
            for bar in bars:
                h = bar.get_height()
                ax.annotate(f"{h:.1f}%", xy=(bar.get_x() + bar.get_width() / 2, h), xytext=(0, 3), textcoords="offset points", ha="center", va="bottom", color="white", fontsize=9, fontweight="bold")
        ax.set_ylim(max(45, min(cnn_f1 + vit_f1 + swin_f1) - 10), 107)
        ax.set_xticks(x)
        ax.set_xticklabels(["Beyin MR\n(Tümör — 4 Sınıf)", "Göğüs Röntgeni\n(COVID/Pnömoni/TB — 4 Sınıf)", "Akciğer BT\n(Kanser — 4 Sınıf)"], color="white", fontsize=10, fontweight="bold")
        ax.set_ylabel("Test Macro F1-Skoru (%)", color="#cbd5e1")
        ax.set_title("3 Medikal Modalitede CNN vs ViT vs Swin Transformer Başarısı", color="#38bdf8", fontsize=13, fontweight="bold", pad=12)
        ax.tick_params(colors="#cbd5e1")
        ax.grid(axis="y", alpha=0.22, color="#64748b")
        ax.legend(facecolor="#0f172a", edgecolor="#334155", labelcolor="white", loc="upper center", ncol=3, fontsize=9)
        plt.tight_layout()
        fig.savefig(plots_dir / "multimodal_overview.png", facecolor=fig.get_facecolor(), bbox_inches="tight")
        plt.close(fig)

    # 2. 3 Modaliteli Grad-CAM + MobileSAM Vitrin Görseli
    showcase_items = [
        ("brain_mri", "glioma", PROJECT_ROOT / "sample_images" / "glioma_01.jpg", True),
        ("chest_xray", "covid19", PROJECT_ROOT / "sample_images" / "chest_xray" / "covid19_01.jpg", False),
        ("chest_ct", "adenocarcinoma", PROJECT_ROOT / "sample_images" / "chest_ct" / "adenocarcinoma_01.jpg", True),
    ]
    valid = [it for it in showcase_items if it[2].exists()]
    if not valid:
        return

    transform = get_eval_transforms(224)
    segmenter = TumorSegmenter()
    fig, axes = plt.subplots(len(valid), 4, figsize=(13.2, 3.4 * len(valid)), dpi=180)
    fig.patch.set_facecolor("#0f172a")
    col_titles = ["Orijinal Medikal Görüntü", "Klasik Ön İşleme (CLAHE)", "Grad-CAM (ResNet-50)", "MobileSAM Lezyon Maskesi"]

    for r_idx, (mod, true_cls, img_p, use_roi) in enumerate(valid):
        m_cfg = cfg["modalities"][mod]
        steps = preprocess_pipeline(img_p, use_roi=use_roi)
        inp = transform(Image.fromarray(steps["final_rgb"])).unsqueeze(0).to(device)

        best_m_data = json.load(open(PROJECT_ROOT / "results" / m_cfg["metrics_file"], encoding="utf-8"))
        best_mk = best_m_data["best_model_key"]
        pred_model, payload = load_trained_model(best_mk, modality=mod, device=device)
        classes = payload["classes"]
        with torch.no_grad():
            probs = torch.softmax(pred_model(inp), dim=1)[0].cpu().numpy()
            pred_idx = int(np.argmax(probs))

        cam_model, _ = load_trained_model("cnn", modality=mod, device=device)
        gc = UnifiedGradCAM(cam_model, "cnn")
        cam_2d, _, _ = gc.generate(inp, target_class=pred_idx)
        gc.remove_hooks()

        if use_roi:
            clean_cam = mask_cam_to_brain(cam_2d, steps["roi_rgb_resized"], erode_px=12)
        else:
            clean_cam = cam_2d

        cam_ov = overlay_gradcam(steps["final_rgb"], clean_cam, alpha=0.45)
        seg = segmenter.segment(steps["roi_rgb_resized"], clean_cam, cam_threshold=0.52, apply_interior_mask=use_roi)

        row_imgs = [steps["original_rgb"], steps["final_rgb"], cam_ov, seg["overlay_rgb"]]
        row_caps = [
            f"{m_cfg['short_name']} | Gerçek: {m_cfg['class_labels_tr'][true_cls]}",
            "ROI + Bilateral + CLAHE" if use_roi else "Bilateral + CLAHE (224x224)",
            f"Tahmin ({cfg['models'][best_mk]['short_name']}): {m_cfg['class_labels_tr'][classes[pred_idx]]} (%{probs[pred_idx]*100:.1f})",
            f"SAM Şüpheli Bölge Oranı: %{seg['tumor_ratio_pct']:.1f}",
        ]
        for c_idx in range(4):
            ax = axes[r_idx, c_idx] if len(valid) > 1 else axes[c_idx]
            ax.imshow(row_imgs[c_idx])
            if r_idx == 0:
                ax.set_title(col_titles[c_idx], color="#38bdf8", fontsize=11, fontweight="bold", pad=8)
            ax.set_xlabel(row_caps[c_idx], color="#e2e8f0", fontsize=8.8)
            ax.set_xticks([])
            ax.set_yticks([])

    fig.suptitle("3 Modaliteli Açıklanabilir Tanı (Grad-CAM) ve MobileSAM Segmentasyonu (MR • Röntgen • BT)", color="white", fontsize=13, fontweight="bold", y=0.99)
    plt.tight_layout(rect=[0, 0, 1, 0.96])
    fig.savefig(plots_dir / "multimodal_gradcam_sam.png", facecolor=fig.get_facecolor(), bbox_inches="tight")
    plt.close(fig)


def main() -> None:
    parser = argparse.ArgumentParser(description="Otomatik Modalite Router, Akciğer BT ve Göğüs Röntgeni modellerini eğitir.")
    parser.add_argument("--targets", nargs="+", default=["router", "chest_ct", "chest_xray"], choices=["router", "chest_ct", "chest_xray", "plots"])
    args = parser.parse_args()

    set_seed(42)
    device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    print(f"Kullanılan Cihaz: {device}" + (f" ({torch.cuda.get_device_name(0)})" if device.type == "cuda" else ""))

    if "router" in args.targets:
        train_router(device, epochs=5)
    if "chest_ct" in args.targets:
        train_modality("chest_ct", device, epochs=12)
    if "chest_xray" in args.targets:
        train_modality("chest_xray", device, epochs=8)

    print("\n3 Modaliteli genel özet ve Grad-CAM/SAM vitrin grafikleri üretiliyor...")
    generate_multimodal_summary_plots(device)
    print("TÜM ÇOKLU MODALİTE EĞİTİMİ VE GRAFİKLER TAMAMLANDI!")


if __name__ == "__main__":
    main()
