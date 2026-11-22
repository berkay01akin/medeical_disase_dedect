"""
Çoklu Modalite Veri Seti Hazırlayıcı (scripts/prepare_multimodal_datasets.py)
-----------------------------------------------------------------------------
Mevcut Beyin MR verisine dokunmadan 2 yeni modalite ve 1 Otomatik Modalite Router
veri setini Hugging Face üzerinden otomatik indirir ve ön-işler:
  1. Akciğer Tomografisi (BT — Kabil007/LungCancer4Types):
     4 Sınıf: adenocarcinoma, large_cell_carcinoma, normal, squamous_cell_carcinoma
     Ön-işleme: Otsu+Morfoloji ROI + Bilateral + CLAHE -> 224x224
  2. Göğüs Röntgeni (X-Ray — fahmiaziz/pneumonia-covid-tuberculosis):
     4 Sınıf: covid19, normal, pneumonia, tuberculosis
     Ön-işleme: Bilateral + CLAHE (ROI kırpmasız — akciğer alanlarını korumak için) -> 224x224
  3. Otomatik Modalite Router (brain_mri / chest_ct / chest_xray):
     Ham (ön-işlenmemiş) 224x224 görüntülerden dengeli train/val/test seti.
"""

from __future__ import annotations

import io
import json
import random
import shutil
import sys
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Dict, List, Tuple

import polars as pl
from huggingface_hub import HfApi, hf_hub_download
from PIL import Image
from tqdm import tqdm

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from src.preprocessing import preprocess_pipeline

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"}


def _normalize_ct_class(folder_name: str) -> str | None:
    s = folder_name.lower().replace("-", "_").replace(".", "_")
    if "adeno" in s:
        return "adenocarcinoma"
    if "large" in s:
        return "large_cell_carcinoma"
    if "squamous" in s:
        return "squamous_cell_carcinoma"
    if "normal" in s:
        return "normal"
    return None


def _process_path_to_jpg(src_path: Path, dst_path: Path, use_roi: bool) -> bool:
    try:
        steps = preprocess_pipeline(
            image_input=src_path,
            use_roi=use_roi,
            filter_type="bilateral",
            kernel_size=5,
            use_clahe=True,
            clahe_clip=2.0,
            clahe_grid=8,
            output_size=224,
        )
        dst_path.parent.mkdir(parents=True, exist_ok=True)
        Image.fromarray(steps["final_rgb"]).save(dst_path, format="JPEG", quality=95)
        return True
    except Exception as exc:
        print(f"[UYARI] {src_path.name} işlenemedi: {exc}")
        return False


def _save_raw_224(src_input, dst_path: Path) -> bool:
    try:
        if isinstance(src_input, (str, Path)):
            img = Image.open(src_input).convert("RGB")
        else:
            img = Image.open(io.BytesIO(src_input)).convert("RGB")
        img = img.resize((224, 224), Image.Resampling.BILINEAR)
        dst_path.parent.mkdir(parents=True, exist_ok=True)
        img.save(dst_path, format="JPEG", quality=92)
        return True
    except Exception:
        return False


def prepare_chest_ct(seed: int = 42) -> Tuple[Dict[str, Dict[str, int]], List[Path]]:
    print("\n" + "=" * 72)
    print("[MODALİTE 2/3] AKCİĞER TOMOGRAFİSİ (BT) — Kabil007/LungCancer4Types")
    print("=" * 72)
    raw_dir = PROJECT_ROOT / "data" / "chest_ct" / "raw"
    proc_dir = PROJECT_ROOT / "data" / "chest_ct" / "processed"
    sample_dir = PROJECT_ROOT / "sample_images" / "chest_ct"

    if not raw_dir.exists() or not any(raw_dir.rglob("*.png")) and not any(raw_dir.rglob("*.jpg")):
        zip_path = hf_hub_download("Kabil007/LungCancer4Types", "Data-20240107T052410Z-001.zip", repo_type="dataset")
        raw_dir.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(raw_dir)
        print(f"  ZIP çıkarıldı -> {raw_dir}")

    # Tüm görüntüleri split (train/valid/test) + sınıf bazlı tara
    by_split: Dict[str, Dict[str, List[Path]]] = {
        "train": {},
        "val": {},
        "test": {},
    }
    all_raw_ct: List[Path] = []
    for p in sorted(raw_dir.rglob("*")):
        if not p.is_file() or p.suffix.lower() not in IMAGE_EXTS:
            continue
        cls = _normalize_ct_class(p.parent.name)
        if cls is None:
            continue
        all_raw_ct.append(p)
        grand = p.parent.parent.name.lower()
        if "test" in grand:
            sp = "test"
        elif "val" in grand:
            sp = "val"
        else:
            sp = "train"
        by_split[sp].setdefault(cls, []).append(p)

    # Eğer val/test boşsa stratified böl
    if not by_split["test"] or not by_split["val"]:
        rng = random.Random(seed)
        merged: Dict[str, List[Path]] = {}
        for sp_dict in by_split.values():
            for c, files in sp_dict.items():
                merged.setdefault(c, []).extend(files)
        by_split = {"train": {}, "val": {}, "test": {}}
        for c, files in merged.items():
            rng.shuffle(files)
            n = len(files)
            n_test = max(1, int(n * 0.20))
            n_val = max(1, int(n * 0.15))
            by_split["test"][c] = files[:n_test]
            by_split["val"][c] = files[n_test : n_test + n_val]
            by_split["train"][c] = files[n_test + n_val :]

    if proc_dir.exists():
        shutil.rmtree(proc_dir)
    if sample_dir.exists():
        shutil.rmtree(sample_dir)
    sample_dir.mkdir(parents=True, exist_ok=True)

    tasks = []
    counts: Dict[str, Dict[str, int]] = {"train": {}, "val": {}, "test": {}}
    for sp, cls_map in by_split.items():
        for cls, files in sorted(cls_map.items()):
            counts[sp][cls] = len(files)
            for idx, src in enumerate(files):
                dst = proc_dir / sp / cls / f"ct_{cls}_{idx:04d}.jpg"
                tasks.append((src, dst))
            if sp == "test":
                for s_idx, s_path in enumerate(files[:2], start=1):
                    Image.open(s_path).convert("RGB").save(sample_dir / f"{cls}_{s_idx:02d}.jpg", quality=95)

    with ThreadPoolExecutor(max_workers=8) as ex:
        futs = [ex.submit(_process_path_to_jpg, s, d, True) for s, d in tasks]
        for _ in tqdm(as_completed(futs), total=len(futs), desc="Akciğer BT Ön İşleme", unit="img"):
            pass

    print("  Akciğer BT dağılımı:")
    for sp in ["train", "val", "test"]:
        print(f"    {sp.upper():5s}: {sum(counts[sp].values()):4d} -> {counts[sp]}")
    return counts, all_raw_ct


def prepare_chest_xray(seed: int = 42, max_per_class: int = 900) -> Tuple[Dict[str, Dict[str, int]], List[bytes]]:
    print("\n" + "=" * 72)
    print("[MODALİTE 3/3] GÖĞÜS RÖNTGENİ (X-RAY) — fahmiaziz/pneumonia-covid-tuberculosis")
    print("=" * 72)
    proc_dir = PROJECT_ROOT / "data" / "chest_xray" / "processed"
    sample_dir = PROJECT_ROOT / "sample_images" / "chest_xray"

    api = HfApi()
    repo_id = "fahmiaziz/pneumonia-covid-tuberculosis"
    parquet_files = sorted([f for f in api.list_repo_files(repo_id, repo_type="dataset") if f.endswith(".parquet")])
    print(f"  {len(parquet_files)} parquet dosyası indiriliyor/okunuyor...")

    label_map = {0: "normal", 1: "pneumonia", 2: "covid19", 3: "tuberculosis"}
    by_class: Dict[str, List[bytes]] = {v: [] for v in label_map.values()}

    for pf in parquet_files:
        local_pf = hf_hub_download(repo_id, pf, repo_type="dataset")
        df = pl.read_parquet(local_pf)
        for row in df.iter_rows(named=True):
            lbl = row["label"]
            cls = label_map.get(int(lbl))
            if cls is None:
                continue
            img_struct = row["image"]
            img_bytes = img_struct["bytes"] if isinstance(img_struct, dict) else img_struct
            if img_bytes:
                by_class[cls].append(img_bytes)

    print("  Ham X-Ray sınıf sayıları:", {k: len(v) for k, v in by_class.items()})

    rng = random.Random(seed)
    if proc_dir.exists():
        shutil.rmtree(proc_dir)
    if sample_dir.exists():
        shutil.rmtree(sample_dir)
    sample_dir.mkdir(parents=True, exist_ok=True)

    counts: Dict[str, Dict[str, int]] = {"train": {}, "val": {}, "test": {}}
    router_raw_xray: List[bytes] = []

    def _proc_bytes(img_bytes: bytes, dst: Path) -> bool:
        try:
            steps = preprocess_pipeline(
                image_input=img_bytes,
                use_roi=False,  # Göğüs röntgeninde akciğer köşelerini korumak için ROI kırpması kapalı
                filter_type="bilateral",
                kernel_size=5,
                use_clahe=True,
                clahe_clip=2.0,
                clahe_grid=8,
                output_size=224,
            )
            dst.parent.mkdir(parents=True, exist_ok=True)
            Image.fromarray(steps["final_rgb"]).save(dst, format="JPEG", quality=95)
            return True
        except Exception:
            return False

    tasks = []
    for cls, items in sorted(by_class.items()):
        rng.shuffle(items)
        selected = items[:max_per_class]
        n = len(selected)
        n_test = int(n * 0.20)
        n_val = int(n * 0.15)
        test_set = selected[:n_test]
        val_set = selected[n_test : n_test + n_val]
        train_set = selected[n_test + n_val :]

        counts["train"][cls] = len(train_set)
        counts["val"][cls] = len(val_set)
        counts["test"][cls] = len(test_set)

        router_raw_xray.extend(selected[:250])

        for sp_name, subset in [("train", train_set), ("val", val_set), ("test", test_set)]:
            for idx, b in enumerate(subset):
                dst = proc_dir / sp_name / cls / f"xray_{cls}_{idx:04d}.jpg"
                tasks.append((b, dst))

        for s_idx, b in enumerate(test_set[:2], start=1):
            Image.open(io.BytesIO(b)).convert("RGB").save(sample_dir / f"{cls}_{s_idx:02d}.jpg", quality=95)

    with ThreadPoolExecutor(max_workers=8) as ex:
        futs = [ex.submit(_proc_bytes, b, d) for b, d in tasks]
        for _ in tqdm(as_completed(futs), total=len(futs), desc="Göğüs Röntgeni Ön İşleme", unit="img"):
            pass

    print("  Göğüs Röntgeni dengeli dağılımı:")
    for sp in ["train", "val", "test"]:
        print(f"    {sp.upper():5s}: {sum(counts[sp].values()):4d} -> {counts[sp]}")
    return counts, router_raw_xray


def prepare_router_dataset(
    raw_ct_paths: List[Path],
    raw_xray_bytes: List[bytes],
    seed: int = 42,
) -> Dict[str, Dict[str, int]]:
    print("\n" + "=" * 72)
    print("[ROUTER] OTOMATİK MODALİTE ALGILAYICI VERİ SETİ (Beyin MR / Göğüs Röntgen / Akciğer BT)")
    print("=" * 72)
    router_dir = PROJECT_ROOT / "data" / "router"
    if router_dir.exists():
        shutil.rmtree(router_dir)

    rng = random.Random(seed)
    brain_raw = PROJECT_ROOT / "data" / "raw"
    brain_paths = [p for p in sorted(brain_raw.rglob("*")) if p.is_file() and p.suffix.lower() in IMAGE_EXTS]

    rng.shuffle(brain_paths)
    rng.shuffle(raw_ct_paths)
    rng.shuffle(raw_xray_bytes)

    per_mod = min(900, len(brain_paths), len(raw_ct_paths), len(raw_xray_bytes))
    pools = {
        "brain_mri": brain_paths[:per_mod],
        "chest_ct": raw_ct_paths[:per_mod],
        "chest_xray": raw_xray_bytes[:per_mod],
    }

    counts: Dict[str, Dict[str, int]] = {"train": {}, "val": {}, "test": {}}
    tasks = []
    for mod, items in pools.items():
        n_test = int(len(items) * 0.15)
        n_val = int(len(items) * 0.15)
        splits = {
            "test": items[:n_test],
            "val": items[n_test : n_test + n_val],
            "train": items[n_test + n_val :],
        }
        for sp, sub in splits.items():
            counts[sp][mod] = len(sub)
            for idx, item in enumerate(sub):
                dst = router_dir / sp / mod / f"{mod}_{idx:04d}.jpg"
                tasks.append((item, dst))

    with ThreadPoolExecutor(max_workers=8) as ex:
        futs = [ex.submit(_save_raw_224, item, d) for item, d in tasks]
        for _ in tqdm(as_completed(futs), total=len(futs), desc="Router Verisi", unit="img"):
            pass

    print("  Router veri dağılımı:")
    for sp in ["train", "val", "test"]:
        print(f"    {sp.upper():5s}: {sum(counts[sp].values()):4d} -> {counts[sp]}")
    return counts


def main() -> None:
    ct_counts, raw_ct_paths = prepare_chest_ct()
    xray_counts, raw_xray_bytes = prepare_chest_xray()
    router_counts = prepare_router_dataset(raw_ct_paths, raw_xray_bytes)

    out_summary = PROJECT_ROOT / "results" / "multimodal_dataset_summary.json"
    out_summary.parent.mkdir(parents=True, exist_ok=True)
    with open(out_summary, "w", encoding="utf-8") as f:
        json.dump(
            {
                "chest_ct": ct_counts,
                "chest_xray": xray_counts,
                "router": router_counts,
            },
            f,
            ensure_ascii=False,
            indent=2,
        )
    print(f"\nTAMAMLANDI! Çoklu modalite veri özeti -> {out_summary}")


if __name__ == "__main__":
    main()
