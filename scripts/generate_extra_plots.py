"""
Ek Görsel Üreticisi (scripts/generate_extra_plots.py)
-----------------------------------------------------
1. results/plots/router_confusion_matrix.png:
   Otomatik Modalite Router (ResNet-18 — 405/405 = %100) karışıklık matrisi grafiği.
2. results/plots/xray_ct_preprocessing_showcase.png:
   Göğüs Röntgeni (Bilateral + CLAHE, ROI Kapalı) ve Akciğer BT (Otsu + Morfoloji ROI + Bilateral + CLAHE)
   klasik görüntü işleme karşılaştırma vitrini.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from src.preprocessing import preprocess_pipeline

PLOTS_DIR = PROJECT_ROOT / "results" / "plots"
SAMPLE_DIR = PROJECT_ROOT / "sample_images"


def make_router_cm() -> None:
    rm = json.load(open(PROJECT_ROOT / "results" / "router_metrics.json", encoding="utf-8"))
    cm = np.array(rm["confusion_matrix"])
    labels = ["Beyin MR", "Akciğer BT", "Göğüs Röntgeni"]

    fig, ax = plt.subplots(figsize=(5.4, 4.4), dpi=180)
    fig.patch.set_facecolor("#0f172a")
    sns.heatmap(
        cm,
        annot=True,
        fmt="d",
        cmap="Greens",
        cbar=False,
        xticklabels=labels,
        yticklabels=labels,
        ax=ax,
        annot_kws={"size": 14, "weight": "bold"},
    )
    ax.set_title(
        f"Otomatik Modalite Router (ResNet-18)\nTest Accuracy: %{rm['accuracy']*100:.2f} (405 / 405)",
        color="#34d399",
        fontsize=11.5,
        fontweight="bold",
        pad=10,
    )
    ax.set_xlabel("Tahmin Edilen Modalite", color="#cbd5e1", fontsize=10)
    ax.set_ylabel("Gerçek Modalite", color="#cbd5e1", fontsize=10)
    ax.tick_params(colors="#f8fafc", labelsize=9.5)
    plt.tight_layout()
    fig.savefig(PLOTS_DIR / "router_confusion_matrix.png", facecolor=fig.get_facecolor(), bbox_inches="tight")
    plt.close(fig)


def make_xray_ct_preprocessing() -> None:
    xray_p = SAMPLE_DIR / "chest_xray" / "covid19_01.jpg"
    ct_p = SAMPLE_DIR / "chest_ct" / "adenocarcinoma_01.jpg"
    if not xray_p.exists() or not ct_p.exists():
        return

    xr = preprocess_pipeline(xray_p, use_roi=False, filter_type="bilateral", kernel_size=5, use_clahe=True, clahe_clip=2.0, clahe_grid=8)
    ct = preprocess_pipeline(ct_p, use_roi=True, filter_type="bilateral", kernel_size=5, use_clahe=True, clahe_clip=2.0, clahe_grid=8)

    fig, axes = plt.subplots(2, 4, figsize=(13.2, 6.4), dpi=180)
    fig.patch.set_facecolor("#0f172a")

    row1 = [
        (xr["original_rgb"], "1. Orijinal Göğüs Röntgeni", None),
        (xr["gray"], "2. Gri Tonlama (ROI Kapalı*)", "gray"),
        (xr["filtered_gray"], "3. Bilateral Gürültü Filtresi", "gray"),
        (xr["final_rgb"], "4. CLAHE Sonrası (224x224)", None),
    ]
    row2 = [
        (ct["original_rgb"], "1. Orijinal Akciğer BT", None),
        (ct["morph_mask"], "2. Otsu + Morfoloji Maskesi", "gray"),
        (ct["roi_gray"], "3. ROI Kırpılmış Akciğer BT", "gray"),
        (ct["final_rgb"], "4. Bilateral + CLAHE (224x224)", None),
    ]

    for c_idx, (img, title, cmap) in enumerate(row1):
        axes[0, c_idx].imshow(img, cmap=cmap)
        axes[0, c_idx].set_title(title, color="#34d399", fontsize=10.5, fontweight="bold", pad=6)
        axes[0, c_idx].axis("off")

    for c_idx, (img, title, cmap) in enumerate(row2):
        axes[1, c_idx].imshow(img, cmap=cmap)
        axes[1, c_idx].set_title(title, color="#a78bfa", fontsize=10.5, fontweight="bold", pad=6)
        axes[1, c_idx].axis("off")

    fig.suptitle(
        "Modaliteye Duyarlı Klasik Görüntü İşleme: Göğüs Röntgeni (*Akciğer Koruma İçin ROI Kapalı) ve Akciğer BT (ROI Açık)",
        color="#38bdf8",
        fontsize=12.2,
        fontweight="bold",
        y=0.98,
    )
    plt.tight_layout(rect=[0, 0, 1, 0.94])
    fig.savefig(PLOTS_DIR / "xray_ct_preprocessing_showcase.png", facecolor=fig.get_facecolor(), bbox_inches="tight")
    plt.close(fig)


def main() -> None:
    PLOTS_DIR.mkdir(parents=True, exist_ok=True)
    make_router_cm()
    make_xray_ct_preprocessing()
    print("Ek görseller üretildi: router_confusion_matrix.png & xray_ct_preprocessing_showcase.png")


if __name__ == "__main__":
    main()
