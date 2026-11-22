"""
Medikal Görüntülerden Hastalık Tespiti — 3 Modaliteli & Otomatik Router Demo Arayüzü (app.py)
---------------------------------------------------------------------------------------------
Çalıştırma:
    streamlit run app.py

Desteklenen 3 Medikal Modalite + Otomatik Modalite Algılayıcı (ResNet-18 Router):
  1. 🧠 Beyin MR          -> Glioma, Menenjiyom, Hipofiz Tümörü, Tümör Yok
  2. 🫁 Göğüs Röntgeni    -> COVID-19, Pnömoni, Tüberküloz, Normal
  3. 🩻 Akciğer BT (CT)   -> Adenokarsinom, Büyük Hücreli, Skuamöz Hücreli Kanser, Normal
"""

from __future__ import annotations

import json
import time
from pathlib import Path
from typing import Any, Dict, List, Tuple

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st
import torch
from PIL import Image

from src.explain_and_segment import TumorSegmenter, UnifiedGradCAM, mask_cam_to_brain, overlay_gradcam
from src.models import (
    get_checkpoint_dir,
    get_eval_transforms,
    get_modality_config,
    load_config,
    load_modality_router,
    load_trained_model,
)
from src.preprocessing import load_medical_image, preprocess_pipeline

PROJECT_ROOT = Path(__file__).resolve().parent
SAMPLE_DIR = PROJECT_ROOT / "sample_images"
RESULTS_DIR = PROJECT_ROOT / "results"
PLOTS_DIR = RESULTS_DIR / "plots"

st.set_page_config(
    page_title="Medikal Görüntülerden Hastalık Tespiti (MR • Röntgen • BT)",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        .main-header {
            background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 55%, #0c4a6e 100%);
            padding: 1.35rem 1.75rem;
            border-radius: 14px;
            border: 1px solid rgba(56, 189, 248, 0.38);
            margin-bottom: 1.15rem;
        }
        .main-header h1 { color: #f8fafc; font-size: 1.8rem; margin: 0; }
        .main-header p { color: #bae6fd; margin: 0.35rem 0 0 0; font-size: 0.95rem; }
        .result-card {
            background: #111827;
            border: 1px solid #334155;
            border-left: 5px solid #38bdf8;
            border-radius: 10px;
            padding: 1rem 1.2rem;
            margin-bottom: 0.8rem;
        }
        .result-card.alert { border-left-color: #f43f5e; }
        .result-card.safe { border-left-color: #10b981; }
        .router-pill {
            background: rgba(16, 185, 129, 0.14);
            border: 1px solid rgba(16, 185, 129, 0.45);
            border-radius: 10px;
            padding: 0.65rem 1rem;
            color: #a7f3d0;
            margin-bottom: 0.85rem;
            font-size: 0.92rem;
        }
        .badge {
            display: inline-block;
            padding: 0.2rem 0.6rem;
            border-radius: 999px;
            font-size: 0.77rem;
            font-weight: 600;
            background: rgba(56, 189, 248, 0.16);
            color: #38bdf8;
            margin-right: 0.35rem;
            margin-top: 0.25rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_resource
def get_device() -> torch.device:
    return torch.device("cuda" if torch.cuda.is_available() else "cpu")


@st.cache_resource
def get_cached_model(modality: str, model_key: str):
    return load_trained_model(model_key, modality=modality, device=get_device())


@st.cache_resource
def get_cached_router():
    router_ckpt = PROJECT_ROOT / "checkpoints" / "router_best.pth"
    if not router_ckpt.exists():
        return None, None
    return load_modality_router(device=get_device())


@st.cache_resource
def get_cached_segmenter() -> TumorSegmenter:
    return TumorSegmenter()


def get_available_models(cfg: Dict[str, Any], modality: str) -> List[str]:
    ckpt_dir = get_checkpoint_dir(modality)
    return [k for k in cfg["models"].keys() if (ckpt_dir / f"{k}_best.pth").exists()]


def load_json_safe(path: Path) -> Dict[str, Any] | None:
    if not path.exists():
        return None
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def predict_modality(raw_rgb: np.ndarray) -> Tuple[str | None, float, Dict[str, float]]:
    """Otomatik Modalite Router (ResNet-18) ile görüntünün MR / Röntgen / BT olduğunu tahmin eder."""
    router_model, payload = get_cached_router()
    if router_model is None or payload is None:
        return None, 0.0, {}

    device = get_device()
    transform = get_eval_transforms(224)
    inp = transform(Image.fromarray(raw_rgb)).unsqueeze(0).to(device)
    with torch.no_grad():
        probs = torch.softmax(router_model(inp), dim=1)[0].cpu().numpy()
    classes = payload["classes"]
    best_idx = int(np.argmax(probs))
    prob_map = {classes[i]: float(probs[i]) for i in range(len(classes))}
    return classes[best_idx], float(probs[best_idx]), prob_map


def run_inference(
    modality: str,
    final_rgb: np.ndarray,
    selected_mode: str,
    available_keys: List[str],
) -> Dict[str, Any]:
    device = get_device()
    transform = get_eval_transforms(224)
    inp_tensor = transform(Image.fromarray(final_rgb)).unsqueeze(0).to(device)

    t0 = time.perf_counter()
    if selected_mode == "ensemble" and len(available_keys) >= 2:
        prob_list = []
        classes = None
        for mk in available_keys:
            model, payload = get_cached_model(modality, mk)
            classes = payload["classes"]
            with torch.no_grad():
                probs_i = torch.softmax(model(inp_tensor), dim=1)[0].cpu().numpy()
            prob_list.append(probs_i)
        probs = np.mean(prob_list, axis=0)
        pred_idx = int(np.argmax(probs))

        # Keskin uzamsal Grad-CAM için CNN (ResNet-50) kullan, yoksa ilk model
        cam_mk = "cnn" if "cnn" in available_keys else available_keys[0]
        cam_model, _ = get_cached_model(modality, cam_mk)
        gc = UnifiedGradCAM(cam_model, cam_mk)
        cam_2d, _, _ = gc.generate(inp_tensor, target_class=pred_idx)
        gc.remove_hooks()
    else:
        mk = selected_mode if selected_mode in available_keys else available_keys[0]
        model, payload = get_cached_model(modality, mk)
        classes = payload["classes"]
        gc = UnifiedGradCAM(model, mk)
        cam_2d, pred_idx, probs = gc.generate(inp_tensor)
        gc.remove_hooks()

    elapsed_ms = (time.perf_counter() - t0) * 1000.0
    return {
        "classes": classes,
        "pred_idx": pred_idx,
        "pred_class": classes[pred_idx],
        "probs": probs,
        "cam_2d": cam_2d,
        "elapsed_ms": elapsed_ms,
    }


def list_sample_images(cfg: Dict[str, Any], selected_modality_ui: str) -> Dict[str, Tuple[Path, str]]:
    """Seçilen moda göre hazır demo örneklerini döndürür (Otomatik modda 3 modaliteden de örnekler)."""
    out: Dict[str, Tuple[Path, str]] = {}
    mods = ["brain_mri", "chest_xray", "chest_ct"] if selected_modality_ui == "auto" else [selected_modality_ui]
    for mod in mods:
        m_cfg = cfg["modalities"][mod]
        subdir = m_cfg.get("sample_subdir", "")
        sdir = (SAMPLE_DIR / subdir) if subdir else SAMPLE_DIR
        if not sdir.exists():
            continue
        for p in sorted(sdir.iterdir()):
            if p.is_file() and p.suffix.lower() in {".jpg", ".jpeg", ".png", ".bmp"}:
                stem_cls = p.stem.rsplit("_", 1)[0]
                tr_cls = m_cfg["class_labels_tr"].get(stem_cls, stem_cls)
                label = f"{m_cfg['icon']} [{m_cfg['short_name']}] {tr_cls} — {p.name}"
                out[label] = (p, mod)
    return out


def main() -> None:
    cfg = load_config()
    device = get_device()
    router_available = (PROJECT_ROOT / "checkpoints" / "router_best.pth").exists()

    # Üst Başlık
    st.markdown(
        f"""
        <div class="main-header">
            <h1>🏥 {cfg["project_title"]}</h1>
            <p>{cfg["project_subtitle"]}</p>
            <div style="margin-top:0.55rem;">
                <span class="badge">🧠 Beyin MR (Tümör)</span>
                <span class="badge">🫁 Göğüs Röntgeni (COVID-19 / Pnömoni / TB)</span>
                <span class="badge">🩻 Akciğer BT (Kanser)</span>
                <span class="badge">🤖 Otomatik Modalite Router</span>
                <span class="badge">CNN • ViT • Swin • MobileSAM • Grad-CAM</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ==================== KENAR ÇUBUĞU ====================
    with st.sidebar:
        st.header("⚙️ Kontrol Paneli")
        gpu_text = f"GPU: {torch.cuda.get_device_name(0)}" if device.type == "cuda" else "CPU Modu"
        st.caption(f"Donanım: **{gpu_text}**")

        st.subheader("1. Görüntüleme Modalitesi")
        mod_choices = {}
        if router_available:
            mod_choices["🤖 Otomatik Algıla (AI Router)"] = "auto"
        for mk, mv in cfg["modalities"].items():
            if get_available_models(cfg, mk):
                mod_choices[f"{mv['icon']} {mv['display_name']}"] = mk

        chosen_mod_label = st.selectbox("Modalite Seçimi", list(mod_choices.keys()))
        selected_modality_ui = mod_choices[chosen_mod_label]

        st.subheader("2. Medikal Görüntü Seç")
        input_mode = st.radio("Kaynak", ["Hazır Örnek Görüntüler (Demo)", "Kendi Görüntünü Yükle (JPG/PNG/DICOM)"])

        image_source = None
        image_name = ""
        sample_true_mod = None

        if input_mode == "Hazır Örnek Görüntüler (Demo)":
            samples = list_sample_images(cfg, selected_modality_ui)
            if samples:
                chosen_s = st.selectbox("Örnek Seçin", list(samples.keys()))
                image_source, sample_true_mod = samples[chosen_s]
                image_name = image_source.name
            else:
                st.warning("Bu modalite için henüz hazır örnek bulunamadı.")
        else:
            uploaded = st.file_uploader("MR / Röntgen / BT Yükle", type=["jpg", "jpeg", "png", "bmp", "dcm"])
            if uploaded is not None:
                image_source = uploaded.read()
                image_name = uploaded.name

        # Otomatik Modalite Router Çalıştırma (görüntü seçildiyse)
        detected_mod = None
        detected_conf = 0.0
        router_probs: Dict[str, float] = {}
        raw_rgb = None
        if image_source is not None:
            raw_rgb = load_medical_image(image_source, filename=image_name)
            if router_available:
                detected_mod, detected_conf, router_probs = predict_modality(raw_rgb)

        # Aktif modaliteyi belirle
        if selected_modality_ui == "auto":
            active_modality = detected_mod if (detected_mod and get_available_models(cfg, detected_mod)) else "brain_mri"
        else:
            active_modality = selected_modality_ui

        active_m_cfg = get_modality_config(active_modality)
        available_keys = get_available_models(cfg, active_modality)

        st.divider()
        st.subheader("3. Uzman Model Seçimi")
        model_options = {}
        if len(available_keys) >= 2:
            model_options["🔥 3'lü Ensemble (CNN + ViT + Swin)"] = "ensemble"
        for k in available_keys:
            model_options[cfg["models"][k]["display_name"]] = k

        selected_model_key = None
        if model_options:
            chosen_model_label = st.selectbox("Sınıflandırıcı", list(model_options.keys()))
            selected_model_key = model_options[chosen_model_label]
        else:
            st.error(f"{active_m_cfg['short_name']} için eğitilmiş ağırlık bulunamadı.")

        st.divider()
        st.subheader("4. Klasik Ön İşleme Ayarları")
        default_roi = bool(active_m_cfg.get("use_roi_default", True))
        use_roi = st.checkbox("Otsu + Morfoloji ROI Kırpma", value=default_roi)
        filter_type = st.selectbox("Gürültü Filtresi", ["bilateral", "gaussian", "median", "none"], index=0)
        kernel_size = st.slider("Filtre Çekirdek Boyutu", 3, 9, 5, step=2)
        use_clahe = st.checkbox("CLAHE Kontrast İyileştirme", value=True)
        clahe_clip = st.slider("CLAHE Clip Limit", 1.0, 5.0, 2.0, step=0.5)
        clahe_grid = st.slider("CLAHE Grid Boyutu", 4, 16, 8, step=2)

        st.divider()
        st.subheader("5. Grad-CAM & SAM Ayarları")
        cam_alpha = st.slider("Grad-CAM Saydamlık", 0.2, 0.8, 0.45, step=0.05)
        sam_thresh = st.slider("SAM Otomatik Kutu Eşiği", 0.30, 0.80, 0.52, step=0.02)
        force_sam = st.checkbox("Sağlıklı tahminde de SAM çalıştır", value=False)

    # Ön-işleme çalıştır
    steps = None
    if raw_rgb is not None:
        steps = preprocess_pipeline(
            image_input=raw_rgb,
            filename=image_name,
            use_roi=use_roi,
            filter_type=filter_type,
            kernel_size=kernel_size,
            use_clahe=use_clahe,
            clahe_clip=clahe_clip,
            clahe_grid=clahe_grid,
            output_size=224,
        )

    # ==================== 4 ANA SEKME ====================
    tab1, tab2, tab3, tab4 = st.tabs([
        "🔬 Canlı Tanı & SAM Segmentasyon",
        "🧪 Klasik Görüntü İşleme Laboratuvarı",
        "📊 Model Karşılaştırma (MR • Röntgen • BT)",
        "📘 Mimari & Kullanım Kılavuzu",
    ])

    # ==================== TAB 1: CANLI TANI & SAM ====================
    with tab1:
        if steps is None or selected_model_key is None:
            st.info("Analiz başlatmak için sol menüden bir MR, Röntgen veya BT görüntüsü seçin/yükleyin.")
        else:
            # Otomatik Modalite Router Bilgi / Yanlış Görüntü Uyarı Bandı
            if detected_mod is not None:
                det_cfg = cfg["modalities"][detected_mod]
                if selected_modality_ui == "auto":
                    st.markdown(
                        f"""
                        <div class="router-pill">
                            🤖 <strong>Otomatik Modalite Router (ResNet-18):</strong> Yüklenen görüntü
                            <strong>{det_cfg['icon']} {det_cfg['display_name']}</strong> olarak algılandı
                            (<strong>%{detected_conf*100:.2f}</strong> güven) ve ilgili uzman modellere yönlendirildi.
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                elif detected_mod != selected_modality_ui and detected_conf >= 0.70:
                    st.warning(
                        f"⚠️ **Modalite Uyuşmazlığı Koruması:** Şu an **{active_m_cfg['short_name']}** seçili ancak "
                        f"AI Router yüklenen görüntüyü **%{detected_conf*100:.1f}** güvenle **{det_cfg['icon']} {det_cfg['short_name']}** "
                        f"olarak algıladı! Doğru klinik tanı için sol menüden **🤖 Otomatik Algıla** veya **{det_cfg['short_name']}** seçmeniz önerilir."
                    )

            with st.spinner(f"{active_m_cfg['short_name']} uzman analizi, Grad-CAM ve MobileSAM çalışıyor..."):
                inf = run_inference(active_modality, steps["final_rgb"], selected_model_key, available_keys)
                pred_cls = inf["pred_class"]
                pred_tr = active_m_cfg["class_labels_tr"].get(pred_cls, pred_cls)
                pred_desc = active_m_cfg["class_descriptions_tr"].get(pred_cls, "")
                confidence = float(inf["probs"][inf["pred_idx"]]) * 100.0
                healthy_cls = active_m_cfg["healthy_class"]

                apply_interior = bool(active_m_cfg.get("use_roi_default", True))
                clean_cam = (
                    mask_cam_to_brain(inf["cam_2d"], steps["roi_rgb_resized"], erode_px=12)
                    if apply_interior else inf["cam_2d"]
                )
                cam_overlay = overlay_gradcam(steps["final_rgb"], clean_cam, alpha=cam_alpha)

                run_seg = (pred_cls != healthy_cls) or force_sam
                seg_result = None
                if run_seg:
                    segmenter = get_cached_segmenter()
                    seg_result = segmenter.segment(
                        steps["roi_rgb_resized"],
                        clean_cam,
                        cam_threshold=sam_thresh,
                        apply_interior_mask=apply_interior,
                    )

            card_cls = "safe" if pred_cls == healthy_cls else "alert"
            st.markdown(
                f"""
                <div class="result-card {card_cls}">
                    <h3 style="margin:0;color:#f8fafc;">{active_m_cfg['icon']} {active_m_cfg['short_name']} Sonucu: {pred_tr} (%{confidence:.2f} Güven)</h3>
                    <p style="margin:0.35rem 0 0 0;color:#cbd5e1;">{pred_desc}</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Aktif Modalite", active_m_cfg["short_name"])
            m2.metric("Tespit Edilen Sınıf", pred_tr.split(" (")[0])
            m3.metric("Güven / Hız", f"%{confidence:.1f} | {inf['elapsed_ms']:.1f} ms")
            if seg_result is not None:
                m4.metric("SAM Lezyon Alan Oranı", f"%{seg_result['tumor_ratio_pct']:.2f}")
            else:
                m4.metric("SAM Lezyon Alan Oranı", "Bulgu Yok")

            st.subheader("Görsel Analiz Akışı")
            c1, c2, c3, c4 = st.columns(4)
            c1.image(steps["original_rgb"], caption=f"1. Orijinal ({image_name})", use_container_width=True)
            c2.image(steps["final_rgb"], caption="2. Klasik Ön İşleme (CLAHE)", use_container_width=True)
            c3.image(cam_overlay, caption="3. Grad-CAM Karar Isı Haritası", use_container_width=True)
            if seg_result is not None:
                c4.image(
                    seg_result["overlay_rgb"],
                    caption=f"4. {seg_result['method_used']} ({seg_result['tumor_pixels']} px)",
                    use_container_width=True,
                )
            else:
                c4.image(steps["roi_rgb_resized"], caption="4. Sağlıklı Bulgu (Segmentasyon Atlandı)", use_container_width=True)

            col_p1, col_p2 = st.columns([2, 1])
            with col_p1:
                st.subheader("Hastalık Sınıf Olasılıkları (%)")
                prob_df = pd.DataFrame({
                    "Sınıf": [active_m_cfg["class_labels_tr"].get(c, c) for c in inf["classes"]],
                    "Olasılık (%)": [round(float(p) * 100.0, 2) for p in inf["probs"]],
                }).sort_values("Olasılık (%)", ascending=False)
                st.bar_chart(prob_df.set_index("Sınıf"), height=230)
            with col_p2:
                if router_probs:
                    st.subheader("AI Router Modalite Skorları")
                    r_df = pd.DataFrame({
                        "Modalite": [cfg["modalities"][k]["short_name"] for k in router_probs.keys()],
                        "Güven (%)": [round(v * 100.0, 2) for v in router_probs.values()],
                    }).sort_values("Güven (%)", ascending=False)
                    st.dataframe(r_df, use_container_width=True, hide_index=True)

            st.caption("⚠️ Bu uygulama yalnızca akademik/eğitim amaçlıdır; kesin tıbbi tanı yerine geçmez.")

    # ==================== TAB 2: KLASİK GÖRÜNTÜ İŞLEME LABORATUVARI ====================
    with tab2:
        if steps is None:
            st.info("Görüntü işleme adımlarını incelemek için sol menüden bir görüntü seçin.")
        else:
            st.subheader(f"Adım Adım Klasik Görüntü İşleme Hattı ({active_m_cfg['short_name']})")
            r1c1, r1c2, r1c3 = st.columns(3)
            r1c1.image(steps["gray"], caption="1. Gri Tonlama (Grayscale)", use_container_width=True)
            r1c2.image(steps["otsu_mask"], caption="2. Otsu Eşikleme Maskesi", use_container_width=True)
            r1c3.image(steps["morph_mask"], caption="3. Morfolojik Kapama + Açma", use_container_width=True)

            r2c1, r2c2, r2c3 = st.columns(3)
            r2c1.image(steps["roi_gray"], caption=f"4. ROI Bölgesi {steps['roi_bbox']}", use_container_width=True)
            r2c2.image(steps["filtered_gray"], caption=f"5. Gürültü Filtresi ({filter_type.upper()}, k={kernel_size})", use_container_width=True)
            r2c3.image(steps["enhanced_gray"], caption=f"6. CLAHE Sonrası (clip={clahe_clip}, grid={clahe_grid})", use_container_width=True)

            st.subheader("Piksel Yoğunluk Histogramı (Öncesi vs Sonrası)")
            fig, ax = plt.subplots(figsize=(10, 3.5))
            fig.patch.set_facecolor("#0f172a")
            ax.set_facecolor("#1e293b")
            bins = np.arange(256)
            ax.plot(bins, steps["hist_before"], color="#94a3b8", linewidth=1.6, label="CLAHE Öncesi")
            ax.fill_between(bins, steps["hist_before"], color="#94a3b8", alpha=0.2)
            ax.plot(bins, steps["hist_after"], color="#38bdf8", linewidth=1.9, label="CLAHE + Filtre Sonrası")
            ax.fill_between(bins, steps["hist_after"], color="#38bdf8", alpha=0.25)
            ax.set_xlabel("Piksel Parlaklığı (0–255)", color="#cbd5e1")
            ax.set_ylabel("Piksel Sayısı", color="#cbd5e1")
            ax.tick_params(colors="#cbd5e1")
            ax.grid(alpha=0.2, color="#64748b")
            ax.legend(facecolor="#0f172a", edgecolor="#334155", labelcolor="white")
            st.pyplot(fig, use_container_width=True)
            plt.close(fig)

    # ==================== TAB 3: ÇOKLU MODALİTE KARŞILAŞTIRMA ====================
    with tab3:
        st.subheader("3 Modaliteli Genel Performans Özeti (MR • Röntgen • BT)")
        router_m = load_json_safe(RESULTS_DIR / "router_metrics.json")
        if router_m:
            st.success(
                f"🤖 **Otomatik Modalite Router (ResNet-18):** {router_m['test_samples']} bağımsız test görüntüsünde "
                f"**%{router_m['accuracy']*100:.2f} Accuracy** | **%{router_m['f1_macro']*100:.2f} Macro F1** | "
                f"**{router_m['inference_ms_per_img']:.2f} ms/img**"
            )

        if (PLOTS_DIR / "multimodal_overview.png").exists():
            st.image(str(PLOTS_DIR / "multimodal_overview.png"), caption="3 Medikal Modalitede CNN vs ViT vs Swin Macro F1 Karşılaştırması", use_container_width=True)

        st.divider()
        view_mod = st.selectbox(
            "Detaylı Sonuçlarını Görmek İstediğiniz Modalite",
            list(cfg["modalities"].keys()),
            format_func=lambda k: f"{cfg['modalities'][k]['icon']} {cfg['modalities'][k]['display_name']}",
        )
        v_cfg = cfg["modalities"][view_mod]
        v_metrics = load_json_safe(RESULTS_DIR / v_cfg["metrics_file"])
        v_plots = (PLOTS_DIR / v_cfg["plots_subdir"]) if v_cfg["plots_subdir"] else PLOTS_DIR

        if v_metrics is None:
            st.info(f"{v_cfg['display_name']} için metrik dosyası henüz oluşturulmadı.")
        else:
            rows = []
            for mk, m in v_metrics["models"].items():
                rows.append({
                    "Model": m["display_name"],
                    "Aile": m["family"],
                    "Param (M)": m["params_millions"],
                    "Accuracy (%)": round(m["accuracy"] * 100, 2),
                    "Precision (%)": round(m["precision_macro"] * 100, 2),
                    "Recall (%)": round(m["recall_macro"] * 100, 2),
                    "Macro F1 (%)": round(m["f1_macro"] * 100, 2),
                    "ROC-AUC (%)": round(m["roc_auc_macro"] * 100, 2),
                    "Hız (ms/img)": m["inference_ms_per_img"],
                })
            st.dataframe(pd.DataFrame(rows).sort_values("Macro F1 (%)", ascending=False), use_container_width=True, hide_index=True)

            ca, cb = st.columns(2)
            if (v_plots / "model_comparison.png").exists():
                ca.image(str(v_plots / "model_comparison.png"), use_container_width=True)
            if (v_plots / "roc_curves.png").exists():
                cb.image(str(v_plots / "roc_curves.png"), use_container_width=True)
            if (v_plots / "confusion_matrices.png").exists():
                st.image(str(v_plots / "confusion_matrices.png"), use_container_width=True)
            if (v_plots / "training_curves.png").exists():
                st.image(str(v_plots / "training_curves.png"), use_container_width=True)

    # ==================== TAB 4: MİMARİ & KILAVUZ ====================
    with tab4:
        st.subheader("3 Modaliteli Sistem Mimarisi ve Kullanılan Teknolojiler")
        st.markdown(
            f"""
            - **Proje Konusu:** {cfg["project_title"]}
            - **GitHub Reposu:** [{cfg["github_url"]}]({cfg["github_url"]})

            #### Desteklenen Modaliteler ve Uzman Modeller
            1. **🤖 Otomatik Modalite Router (`ResNet-18`):** Yüklenen ham görüntünün Beyin MR, Göğüs Röntgeni veya Akciğer BT olduğunu otomatik algılar ve yanlış modalite yüklemelerini engeller.
            2. **🧠 Beyin MR (4 Sınıf):** Glioma, Menenjiyom, Hipofiz Tümörü, Tümör Yok *(Kaggle 7.200 MR + BRISC 2025 Harici Test)*.
            3. **🫁 Göğüs Röntgeni (4 Sınıf):** COVID-19, Pnömoni, Tüberküloz, Normal *(Hugging Face 6.461 Röntgen)*.
            4. **🩻 Akciğer Tomografisi / BT (4 Sınıf):** Adenokarsinom, Büyük Hücreli Karsinom, Skuamöz Hücreli Karsinom, Normal *(Hugging Face LungCancer4Types)*.

            #### Tüm Modalitelerde Ortak Teknoloji Yığını
            - **Klasik Görüntü İşleme:** OpenCV (Otsu, Morfoloji, Bilateral/Median/Gaussian, CLAHE), `pydicom` (DICOM `.dcm`).
            - **Sınıflandırma:** PyTorch + `timm` (`ResNet-50` CNN, `ViT-B/16`, `Swin-T` + 3'lü Ensemble).
            - **Açıklanabilirlik & Segmentasyon:** Özel **Unified Grad-CAM** + **Ultralytics MobileSAM**.
            """
        )
        if (PLOTS_DIR / "multimodal_gradcam_sam.png").exists():
            st.image(str(PLOTS_DIR / "multimodal_gradcam_sam.png"), caption="3 Modalitede Uçtan Uca Grad-CAM + MobileSAM Çıktısı", use_container_width=True)

        st.subheader("Grup Üyeleri")
        team_df = pd.DataFrame(cfg["team_members"])
        team_df.columns = ["Ad Soyad", "Öğrenci No", "Sorumluluk Alanı"]
        st.dataframe(team_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
