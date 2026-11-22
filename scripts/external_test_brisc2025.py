"""
Harici Klinik Veri Seti Testi — BRISC 2025 (scripts/external_test_brisc2025.py)
------------------------------------------------------------------------------
Kaggle veri setinden tamamen bağımsız olan BRISC 2025 (Hugging Face: chehablab/BRISC2025)
veri setinden yeni MR görüntüleri çıkarır:
  1. external_test_images/       -> Sınıf adıyla etiketli 40 adet yeni MR (app.py ile test için)
  2. external_test_images/kor_test/ -> İsimsiz 16 adet kör test MR + cevap_anahtari.txt
  3. Eğitilen CNN, ViT, Swin ve 3'lü Ensemble modellerinin bu harici veri seti üzerindeki
     genelleme başarısını (Accuracy, Macro F1) ölçer ve results/external_test_results.json kaydeder.
"""

from __future__ import annotations

import io
import json
import random
import sys
from pathlib import Path
from typing import Dict, List

import numpy as np
import polars as pl
import torch
from huggingface_hub import hf_hub_download
from PIL import Image
from sklearn.metrics import accuracy_score, classification_report, f1_score

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from src.models import get_eval_transforms, load_config, load_trained_model
from src.preprocessing import preprocess_pipeline


def infer_class_from_source(source_file: str) -> str | None:
    s = source_file.lower()
    if "_gl_" in s or "glioma" in s:
        return "glioma"
    if "_me_" in s or "meningioma" in s:
        return "meningioma"
    if "_pi_" in s or "pituitary" in s:
        return "pituitary"
    if "_no_" in s or "notumor" in s or "normal" in s:
        return "notumor"
    return None


def main() -> None:
    cfg = load_config()
    classes = cfg["classes"]
    cls_to_idx = {c: i for i, c in enumerate(classes)}

    print("[1/4] BRISC 2025 (chehablab/BRISC2025) parquet dosyası yükleniyor...")
    parquet_path = hf_hub_download("chehablab/BRISC2025", "data/train-00000-of-00001.parquet", repo_type="dataset")
    df = pl.read_parquet(parquet_path)

    # Sınıf eşleştirmesi
    mapped_rows = []
    class_counts: Dict[str, int] = {c: 0 for c in classes}
    for row in df.iter_rows(named=True):
        cls = infer_class_from_source(row["source_file"])
        if cls in cls_to_idx:
            class_counts[cls] += 1
            mapped_rows.append((cls, row["source_file"], row["image"]["bytes"]))

    print(f"Toplam eşleşen BRISC 2025 görüntüsü: {len(mapped_rows)} -> {class_counts}")

    # 2. Harici test klasörüne örnekler kaydet
    ext_dir = PROJECT_ROOT / "external_test_images"
    blind_dir = ext_dir / "kor_test"
    ext_dir.mkdir(parents=True, exist_ok=True)
    blind_dir.mkdir(parents=True, exist_ok=True)

    rng = random.Random(2026)
    by_class: Dict[str, List] = {c: [] for c in classes}
    for item in mapped_rows:
        by_class[item[0]].append(item)

    saved_labeled = 0
    blind_pool = []
    for cls in classes:
        items = by_class[cls]
        rng.shuffle(items)
        # Her sınıftan 10 adet etiketli harici test görüntüsü
        for idx, (c, src_name, img_bytes) in enumerate(items[:10], start=1):
            img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
            out_path = ext_dir / f"brisc2025_{c}_{idx:02d}.jpg"
            img.save(out_path, format="JPEG", quality=95)
            saved_labeled += 1
            if idx <= 4:
                blind_pool.append((c, src_name, img))

    # 16 adet isimsiz kör test görüntüsü + cevap anahtarı
    rng.shuffle(blind_pool)
    answer_lines = ["Dosya Adı | Gerçek Sınıf (Türkçe) | Kaynak ID", "-" * 65]
    for b_idx, (c, src_name, img) in enumerate(blind_pool, start=1):
        fname = f"yeni_test_{b_idx:02d}.jpg"
        img.save(blind_dir / fname, format="JPEG", quality=95)
        answer_lines.append(f"{fname:16s} | {cfg['class_labels_tr'][c]:26s} | {src_name}")
    (blind_dir / "cevap_anahtari.txt").write_text("\n".join(answer_lines), encoding="utf-8")

    print(f"[2/4] {saved_labeled} etiketli + {len(blind_pool)} kör test görüntüsü kaydedildi -> {ext_dir}")

    # 3. Modelleri BRISC 2025 üzerinde değerlendir (hızlı test için her sınıftan max 100 görüntü)
    eval_items = []
    for cls in classes:
        eval_items.extend(by_class[cls][:100])
    print(f"[3/4] {len(eval_items)} adet tamamen harici BRISC 2025 MR üzerinde modeller test ediliyor...")

    device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    transform = get_eval_transforms(224)

    # Tüm görüntüleri bir kez ön-işle
    tensors = []
    y_true = []
    for cls, _, img_bytes in eval_items:
        steps = preprocess_pipeline(img_bytes)
        t = transform(Image.fromarray(steps["final_rgb"]))
        tensors.append(t)
        y_true.append(cls_to_idx[cls])
    x_all = torch.stack(tensors)
    y_true_np = np.array(y_true)

    active_classes = sorted(list(set(y_true)))
    active_names = [classes[i] for i in active_classes]

    results = {}
    prob_bank = []
    for mk in ["cnn", "vit", "swin"]:
        model, _ = load_trained_model(mk, device=device)
        preds_probs = []
        with torch.no_grad():
            for i in range(0, len(x_all), 32):
                batch = x_all[i : i + 32].to(device)
                logits = model(batch)
                preds_probs.append(torch.softmax(logits, dim=1).cpu().numpy())
        probs = np.concatenate(preds_probs, axis=0)
        prob_bank.append(probs)
        y_pred = probs.argmax(axis=1)
        acc = accuracy_score(y_true_np, y_pred)
        f1 = f1_score(y_true_np, y_pred, labels=active_classes, average="macro")
        results[mk] = {
            "display_name": cfg["models"][mk]["display_name"],
            "accuracy": round(float(acc), 4),
            "f1_macro": round(float(f1), 4),
        }
        print(f"  {cfg['models'][mk]['display_name']:32s} -> Acc: {acc*100:5.2f}% | Macro F1: {f1*100:5.2f}%")

    # 3'lü Ensemble
    ens_probs = np.mean(prob_bank, axis=0)
    ens_pred = ens_probs.argmax(axis=1)
    ens_acc = accuracy_score(y_true_np, ens_pred)
    ens_f1 = f1_score(y_true_np, ens_pred, labels=active_classes, average="macro")
    results["ensemble"] = {
        "display_name": "3'lü Ensemble (CNN + ViT + Swin)",
        "accuracy": round(float(ens_acc), 4),
        "f1_macro": round(float(ens_f1), 4),
    }
    print(f"  {'3lü Ensemble (CNN + ViT + Swin)':32s} -> Acc: {ens_acc*100:5.2f}% | Macro F1: {ens_f1*100:5.2f}%")

    print("\nEnsemble Sınıf Bazlı Rapor (BRISC 2025 Harici Veri):")
    print(classification_report(y_true_np, ens_pred, labels=active_classes, target_names=active_names, digits=4))

    out_json = PROJECT_ROOT / "results" / "external_test_results.json"
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(
            {
                "dataset": "chehablab/BRISC2025 (Harici Klinik MR Testi)",
                "evaluated_images": len(eval_items),
                "class_counts_available": class_counts,
                "results": results,
            },
            f,
            ensure_ascii=False,
            indent=2,
        )
    print(f"[4/4] Harici test sonuçları kaydedildi -> {out_json}")


if __name__ == "__main__":
    main()
