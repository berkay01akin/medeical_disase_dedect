"""
Veri Seti Hazırlama ve Klasik Görüntü İşleme Ön-Hazırlık Betiği (scripts/prepare_dataset.py)
--------------------------------------------------------------------------------------------
1. Kaggle'dan indirilen 'brain tumor.zip' arşivini açar.
2. Training klasörünü stratified (%85 train / %15 val) olarak böler; Testing klasörünü test olarak korur.
3. Her görüntüye standart klasik görüntü işleme hattımızı uygular:
   Otsu + Morfolojik ROI Kırpma -> Bilateral Gürültü Filtresi -> CLAHE Kontrast İyileştirme -> 224x224
4. İşlenmiş görüntüleri hızlı GPU eğitimi için data/processed/{train,val,test} altına kaydeder.
5. Canlı sunum demosu için her sınıftan 3 adet orijinal test MR görüntüsünü sample_images/ klasörüne kopyalar.
6. Sunum ve rapor için örnek ön-işleme adım görselini (results/plots/preprocessing_showcase.png) üretir.
"""

from __future__ import annotations

import argparse
import json
import random
import shutil
import sys
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import List, Tuple

import cv2
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from PIL import Image
from tqdm import tqdm

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from src.models import load_config
from src.preprocessing import preprocess_for_training, preprocess_pipeline

DEFAULT_ZIP = Path(r"C:\Users\berka\Downloads\brain tumor.zip")
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"}


def process_one_image(task: Tuple[Path, Path]) -> bool:
    src_path, dst_path = task
    try:
        processed_rgb = preprocess_for_training(src_path, output_size=224)
        dst_path.parent.mkdir(parents=True, exist_ok=True)
        Image.fromarray(processed_rgb).save(dst_path, format="JPEG", quality=95)
        return True
    except Exception as exc:
        print(f"[UYARI] {src_path.name} işlenemedi: {exc}")
        return False


def save_preprocessing_showcase(sample_image_path: Path, output_path: Path) -> None:
    """Rapor ve sunum için klasik görüntü işleme adımlarını gösteren 6 panelli görsel üretir."""
    steps = preprocess_pipeline(sample_image_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    fig, axes = plt.subplots(2, 3, figsize=(13, 8), dpi=180)
    fig.patch.set_facecolor("#0f172a")

    panels = [
        (steps["original_rgb"], "1. Orijinal Beyin MR", None),
        (steps["otsu_mask"], "2. Otsu Eşikleme Maskesi", "gray"),
        (steps["morph_mask"], "3. Morfolojik Kapama + Açma", "gray"),
        (steps["roi_gray"], "4. ROI Kırpılmış Beyin", "gray"),
        (steps["filtered_gray"], "5. Bilateral Gürültü Filtresi", "gray"),
        (steps["final_rgb"], "6. CLAHE Kontrast İyileştirme (224x224)", None),
    ]

    for ax, (img, title, cmap) in zip(axes.flat, panels):
        ax.imshow(img, cmap=cmap)
        ax.set_title(title, color="white", fontsize=11, fontweight="bold", pad=8)
        ax.axis("off")

    fig.suptitle(
        "Klasik Görüntü İşleme Ön İşleme Hattı (Otsu + Morfoloji + Bilateral + CLAHE)",
        color="#38bdf8",
        fontsize=14,
        fontweight="bold",
        y=0.98,
    )
    plt.tight_layout(rect=[0, 0, 1, 0.94])
    fig.savefig(output_path, facecolor=fig.get_facecolor(), bbox_inches="tight")
    plt.close(fig)

    # Histogram karşılaştırma grafiği
    hist_path = output_path.parent / "clahe_histogram_comparison.png"
    fig2, ax2 = plt.subplots(figsize=(9, 4), dpi=180)
    fig2.patch.set_facecolor("#0f172a")
    ax2.set_facecolor("#1e293b")
    bins = np.arange(256)
    ax2.plot(bins, steps["hist_before"], color="#94a3b8", linewidth=1.8, label="CLAHE Öncesi (ROI)")
    ax2.fill_between(bins, steps["hist_before"], color="#94a3b8", alpha=0.2)
    ax2.plot(bins, steps["hist_after"], color="#38bdf8", linewidth=2.0, label="CLAHE Sonrası (İyileştirilmiş)")
    ax2.fill_between(bins, steps["hist_after"], color="#38bdf8", alpha=0.25)
    ax2.set_title("CLAHE Öncesi ve Sonrası Piksel Yoğunluk Histogramı", color="white", fontsize=12, fontweight="bold")
    ax2.set_xlabel("Piksel Parlaklık Değeri (0 - 255)", color="#cbd5e1")
    ax2.set_ylabel("Piksel Sayısı", color="#cbd5e1")
    ax2.tick_params(colors="#cbd5e1")
    ax2.grid(alpha=0.2, color="#64748b")
    ax2.legend(facecolor="#0f172a", edgecolor="#334155", labelcolor="white")
    plt.tight_layout()
    fig2.savefig(hist_path, facecolor=fig2.get_facecolor(), bbox_inches="tight")
    plt.close(fig2)


def main() -> None:
    parser = argparse.ArgumentParser(description="Brain Tumor MRI veri setini çıkarır ve ön-işler.")
    parser.add_argument("--zip-path", type=Path, default=DEFAULT_ZIP, help="Kaggle ZIP dosyasının yolu")
    parser.add_argument("--val-ratio", type=float, default=0.15, help="Doğrulama (validation) oranı")
    parser.add_argument("--seed", type=int, default=42, help="Rastgelelik tohumu")
    parser.add_argument("--workers", type=int, default=8, help="Paralel ön-işleme iş parçacığı sayısı")
    args = parser.parse_args()

    cfg = load_config()
    classes = cfg["classes"]

    raw_dir = PROJECT_ROOT / "data" / "raw"
    proc_dir = PROJECT_ROOT / "data" / "processed"
    sample_dir = PROJECT_ROOT / "sample_images"
    plots_dir = PROJECT_ROOT / "results" / "plots"

    # 1. ZIP dosyasını çıkar
    if not (raw_dir / "Training").exists():
        if not args.zip_path.exists():
            raise FileNotFoundError(f"ZIP dosyası bulunamadı: {args.zip_path}")
        print(f"[1/4] ZIP arşivi çıkarılıyor: {args.zip_path} -> {raw_dir}")
        raw_dir.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(args.zip_path, "r") as zf:
            zf.extractall(raw_dir)
    else:
        print(f"[1/4] Ham veri zaten çıkarılmış: {raw_dir}")

    # 2. Dosyaları topla ve train/val/test böl
    rng = random.Random(args.seed)
    tasks: List[Tuple[Path, Path]] = []
    split_counts = {"train": {}, "val": {}, "test": {}}

    if sample_dir.exists():
        shutil.rmtree(sample_dir)
    sample_dir.mkdir(parents=True, exist_ok=True)

    showcase_src: Path | None = None

    for cls in classes:
        train_files = sorted([p for p in (raw_dir / "Training" / cls).iterdir() if p.suffix.lower() in IMAGE_EXTS])
        test_files = sorted([p for p in (raw_dir / "Testing" / cls).iterdir() if p.suffix.lower() in IMAGE_EXTS])

        rng.shuffle(train_files)
        val_n = int(len(train_files) * args.val_ratio)
        val_split = train_files[:val_n]
        train_split = train_files[val_n:]

        split_counts["train"][cls] = len(train_split)
        split_counts["val"][cls] = len(val_split)
        split_counts["test"][cls] = len(test_files)

        for split_name, files in [("train", train_split), ("val", val_split), ("test", test_files)]:
            for idx, src in enumerate(files):
                dst = proc_dir / split_name / cls / f"{cls}_{idx:04d}.jpg"
                tasks.append((src, dst))

        # Her sınıftan 3 adet orijinal test MR görüntüsünü sample_images/ içine al
        chosen_samples = test_files[:3]
        for s_idx, s_path in enumerate(chosen_samples, start=1):
            dst_sample = sample_dir / f"{cls}_{s_idx:02d}{s_path.suffix.lower()}"
            shutil.copy2(s_path, dst_sample)
            if cls == "glioma" and s_idx == 1:
                showcase_src = dst_sample

    # 3. Klasik Görüntü İşleme hattını paralel çalıştır
    print(f"[2/4] Toplam {len(tasks)} MR görüntüsüne ROI + Bilateral + CLAHE uygulanıyor...")
    ok_count = 0
    with ThreadPoolExecutor(max_workers=args.workers) as executor:
        futures = [executor.submit(process_one_image, t) for t in tasks]
        for fut in tqdm(as_completed(futures), total=len(futures), desc="Ön İşleme", unit="img"):
            if fut.result():
                ok_count += 1

    print(f"[3/4] {ok_count}/{len(tasks)} görüntü başarıyla işlendi -> {proc_dir}")

    # 4. Veri seti özeti ve ön-işleme görseli
    summary = {
        "total_images": ok_count,
        "splits": split_counts,
        "preprocessing": {
            "roi_cropping": "GaussianBlur + Otsu Threshold + Morphological Close/Open + Largest Contour",
            "noise_filter": "Bilateral Filter (d=5, sigmaColor=50, sigmaSpace=50)",
            "contrast_enhancement": "CLAHE (clipLimit=2.0, tileGridSize=8x8)",
            "output_size": "224x224 RGB",
        },
    }
    (PROJECT_ROOT / "results").mkdir(parents=True, exist_ok=True)
    with open(PROJECT_ROOT / "results" / "dataset_summary.json", "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    if showcase_src is not None:
        print(f"[4/4] Ön-işleme vitrin görseli kaydediliyor...")
        save_preprocessing_showcase(showcase_src, plots_dir / "preprocessing_showcase.png")

    print("\nVeri Seti Dağılımı:")
    for sp in ["train", "val", "test"]:
        total_sp = sum(split_counts[sp].values())
        print(f"  {sp.upper():5s}: {total_sp:4d} -> {split_counts[sp]}")
    print(f"Canlı demo örnek görüntüleri: {sample_dir} ({len(list(sample_dir.iterdir()))} dosya)")


if __name__ == "__main__":
    main()
