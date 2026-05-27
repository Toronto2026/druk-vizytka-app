"""
Streamlit-додаток: Агент друку дипломів та подяк
Фестиваль Toronto (toronto.org.ua) → Типографія Визитка
"""

import io
import os
import sys
import tempfile
import traceback
from contextlib import redirect_stdout
from datetime import datetime

import streamlit as st

# ── page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Агент друку — Toronto",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS: скрываем «Made with Streamlit», делаем таблицы шире ──────────────────
st.markdown("""
<style>
footer {visibility: hidden;}
.block-container {padding-top: 1.5rem;}
[data-testid="stDataFrame"] {width: 100%;}
</style>
""", unsafe_allow_html=True)

# ── sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.image(
        "https://toronto.org.ua/wp-content/uploads/2022/09/toronto-logo.png",
        width=180,
    )
    st.markdown("### ⚙️ Налаштування")

    default_month = datetime.now().strftime("%B %Y")
    month = st.text_input("Місяць", value="Січень 2026",
                          help="Відображається у заголовку таблиці")

    st.markdown("---")
    st.markdown("**Бітрікс REST API** *(опційно)*")
    bitrix_url = st.text_input(
        "Webhook URL",
        type="password",
        placeholder="https://your.bitrix24.ua/rest/1/xxx/",
        help="Залиште порожнім, щоб пропустити запис у CRM",
    )
    do_bitrix = st.toggle("Записати номери у Бітрікс після обробки",
                          value=False, disabled=not bitrix_url)

    st.markdown("---")
    st.markdown(
        "<small>ТЗ v8.1 · Excel-пріоритет · [agent_druk.py](https://github.com)</small>",
        unsafe_allow_html=True,
    )

# ── header ─────────────────────────────────────────────────────────────────────
st.title("🎓 Агент друку дипломів та подяк")
st.caption("Фестиваль Toronto · Типографія Визитка")

# ── file uploaders ─────────────────────────────────────────────────────────────
st.markdown("### 1 · Завантажте файли")
col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("**📊 Excel з Бітрікс**")
    excel_file = st.file_uploader(
        "Excel-файл угод",
        type=["xlsx"],
        label_visibility="collapsed",
        help="Аркуші «Друк дипломів» + «Друк подяк», або головний аркуш — агент відфільтрує сам",
    )
    if excel_file:
        st.success(f"✔ {excel_file.name} ({excel_file.size // 1024} KB)")

with col2:
    st.markdown("**📄 PDF № дипломів**")
    pdf_diplomy = st.file_uploader(
        "PDF дипломів",
        type=["pdf"],
        label_visibility="collapsed",
        help="«ТАБЛИЦЯ ДЛЯ ШВИДКОГО ПОШУКУ № ДИПЛОМА»",
    )
    if pdf_diplomy:
        st.success(f"✔ {pdf_diplomy.name} ({pdf_diplomy.size // 1024} KB)")

with col3:
    st.markdown("**📄 PDF № подяк**")
    pdf_podyaky = st.file_uploader(
        "PDF подяк",
        type=["pdf"],
        label_visibility="collapsed",
        help="«ТАБЛИЦЯ ДЛЯ ШВИДКОГО ПОШУКУ № ПОДЯКИ»",
    )
    if pdf_podyaky:
        st.success(f"✔ {pdf_podyaky.name} ({pdf_podyaky.size // 1024} KB)")

all_uploaded = excel_file and pdf_diplomy and pdf_podyaky

with st.expander("📝 Редакція подяк *(опційно)*", expanded=False):
    st.caption(
        "Завантажте Excel-редакцію подяк (колонки: ПІБ керівника | School of arts | Учасники | № подяки). "
        "Агент перевірить, які з них є у списку друку — порівняння за ПІБ керівника (fuzzy match)."
    )
    redaktsiia_file = st.file_uploader(
        "Excel редакції подяк",
        type=["xlsx"],
        label_visibility="collapsed",
        key="redaktsiia_uploader",
        help="Наприклад: zavdannia-dizaineru-v3-with-ids.xlsx",
    )
    if redaktsiia_file:
        st.success(f"✔ {redaktsiia_file.name} ({redaktsiia_file.size // 1024} KB)")

st.markdown("---")
st.markdown("### 2 · Запустити")

run_btn = st.button(
    "▶ Запустити агент",
    type="primary",
    disabled=not all_uploaded,
    use_container_width=False,
)

if not all_uploaded:
    st.info("Завантажте всі три файли, щоб активувати кнопку.")

# ── run ─────────────────────────────────────────────────────────────────────────
if run_btn and all_uploaded:

    # Зберігаємо завантажені файли у тимчасову папку
    with tempfile.TemporaryDirectory() as tmpdir:
        excel_path       = os.path.join(tmpdir, "input.xlsx")
        pdfd_path        = os.path.join(tmpdir, "diplomy.pdf")
        podyp_path       = os.path.join(tmpdir, "podyaky.pdf")
        redaktsiia_path  = os.path.join(tmpdir, "redaktsiia.xlsx")
        output_path      = os.path.join(tmpdir, "output.xlsx")

        with open(excel_path, "wb") as f:
            f.write(excel_file.getbuffer())
        with open(pdfd_path, "wb") as f:
            f.write(pdf_diplomy.getbuffer())
        with open(podyp_path, "wb") as f:
            f.write(pdf_podyaky.getbuffer())
        if redaktsiia_file:
            with open(redaktsiia_path, "wb") as f:
                f.write(redaktsiia_file.getbuffer())

        # Імпортуємо агент (після запису файлів, щоб уникнути проблем)
        try:
            from agent_druk import (
                DEFAULT_CONFIG,
                build_zvedena,
                fuzzy_match,
                process_diplomy,
                process_podyaky,
                read_excel,
                read_all_rows,
                read_pdf_diplomy,
                read_pdf_podyaky,
                read_redaktsiia_podyaky,
                update_bitrix_all,
                write_output,
            )
        except ImportError as e:
            st.error(f"❌ Не вдалося імпортувати agent_druk: {e}")
            st.stop()

        config = {
            **DEFAULT_CONFIG,
            "BITRIX_WEBHOOK_URL": bitrix_url or "",
        }

        errors: list = []
        log_buf = io.StringIO()

        # ── Крок 1: читання ──────────────────────────────────────────────
        with st.status("⏳ Обробка файлів...", expanded=True) as status:

            st.write("📂 Крок 1: Читання файлів...")
            try:
                with redirect_stdout(log_buf):
                    diplomy_rows, podyaky_rows = read_excel(excel_path, config)
                    diplomy_pdf = read_pdf_diplomy(pdfd_path)
                    podyaky_pdf = read_pdf_podyaky(podyp_path)
            except Exception as e:
                status.update(label="❌ Помилка читання", state="error")
                st.error(str(e))
                st.code(traceback.format_exc())
                st.stop()

            # ── Крок 2 ───────────────────────────────────────────────────
            st.write(f"⚙️ Крок 2: Обробка дипломів ({len(diplomy_rows)} рядків)...")
            with redirect_stdout(log_buf):
                diploma_out, podyaka_from_d = process_diplomy(
                    diplomy_rows, diplomy_pdf, podyaky_pdf, config, errors
                )

            # ── Крок 3 ───────────────────────────────────────────────────
            st.write(f"⚙️ Крок 3: Обробка подяк ({len(podyaky_rows)} рядків)...")
            with redirect_stdout(log_buf):
                podyaka_from_p = process_podyaky(
                    podyaky_rows, podyaky_pdf, config, errors
                )

            podyaka_out_all = sorted(
                podyaka_from_d + podyaka_from_p,
                key=lambda r: r["id"],
                reverse=True,
            )

            # ── Крок 4 ───────────────────────────────────────────────────
            st.write("📋 Крок 4: Формування зведеної таблиці...")
            with redirect_stdout(log_buf):
                zvedena = build_zvedena(diploma_out, podyaka_out_all, diplomy_rows)
                write_output(diploma_out, podyaka_out_all, zvedena,
                             output_path, month, errors)

            # ── Крок 4б (опційно) — редакція подяк ──────────────────────
            redaktsiia_data = []
            if redaktsiia_file:
                st.write("📝 Крок 4б: Перевірка редакції подяк...")
                with redirect_stdout(log_buf):
                    redaktsiia_data = read_redaktsiia_podyaky(redaktsiia_path)
                # Звіряємо за ПІБ керівника (fuzzy match, поріг 0.75)
                # № подяки у редакції (1–61) та у списку друку (91, 43, 200…)
                # — різні системи нумерації, тому порівнюємо тільки за ПІБ
                print_recs = [r for r in podyaka_out_all if not r.get("warning")]
                for rec in redaktsiia_data:
                    matched = None
                    pib_red = rec["pib_kerivnyk"]
                    for pr in print_recs:
                        if fuzzy_match(pib_red, pr["pib"], 0.75):
                            matched = pr
                            break
                    rec["in_druk"]   = matched is not None
                    rec["druk_num"]  = matched["num_doc"] if matched else None
                    rec["druk_pib"]  = matched["pib"]    if matched else None

            # ── Крок 5 (опційно) — записати для ВСІХ угод ───────────────
            if do_bitrix and bitrix_url:
                st.write("🔗 Крок 5: Запис №Диплома і №Подяки у Бітрікс для ВСІХ угод...")
                with redirect_stdout(log_buf):
                    all_rows = read_all_rows(excel_path)
                    update_bitrix_all(all_rows, diplomy_pdf, podyaky_pdf, config, errors)

            status.update(label="✅ Готово!", state="complete", expanded=False)

        # ── Зберігаємо вивід у session_state ──────────────────────────────
        with open(output_path, "rb") as f:
            xlsx_bytes = f.read()

        st.session_state["result"] = {
            "xlsx_bytes":    xlsx_bytes,
            "diploma_out":   diploma_out,
            "podyaka_out":   podyaka_out_all,
            "zvedena":       zvedena,
            "errors":        errors,
            "log":           log_buf.getvalue(),
            "month":         month,
            "n_diplomy_pdf": len(diplomy_pdf),
            "n_podyaky_pdf": len(podyaky_pdf),
            "redaktsiia":    redaktsiia_data,
        }

# ── Показуємо результат ─────────────────────────────────────────────────────────
if "result" in st.session_state:
    res = st.session_state["result"]
    diploma_out  = res["diploma_out"]
    podyaka_out  = res["podyaka_out"]
    zvedena      = res["zvedena"]
    errors       = res["errors"]

    st.markdown("---")
    st.markdown("### 3 · Результат")

    # ── Метрики ────────────────────────────────────────────────────────────
    good_d = sum(1 for r in diploma_out if not r.get("warning"))
    good_p = sum(1 for r in podyaka_out if not r.get("warning"))
    qty_d  = sum(r["qty"] for r in diploma_out if not r.get("warning"))
    qty_p  = sum(r["qty"] for r in podyaka_out if not r.get("warning"))

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Дипломів (позицій)", good_d,
              delta=f"−{len(diploma_out)-good_d} помилок" if errors else None,
              delta_color="inverse")
    m2.metric("Примірників дипломів", qty_d)
    m3.metric("Подяк (позицій)", good_p)
    m4.metric("Примірників подяк", qty_p)

    # ── Кнопка завантаження ────────────────────────────────────────────────
    fname = f"Друк_Визитка_{res['month'].replace(' ', '_')}.xlsx"
    st.download_button(
        label="⬇ Завантажити Excel для типографії",
        data=res["xlsx_bytes"],
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=False,
    )

    # ── Вкладки з таблицями ────────────────────────────────────────────────
    redaktsiia   = res.get("redaktsiia", [])
    n_in_druk    = sum(1 for r in redaktsiia if r.get("in_druk"))
    redaktsiia_tab_label = f"📝 Редакція ({n_in_druk}/{len(redaktsiia)} у друку)" if redaktsiia else "📝 Редакція"

    tab_z, tab_d, tab_p, tab_r, tab_err, tab_log = st.tabs([
        "📋 Зведена",
        f"🎓 Диплом ({len(diploma_out)})",
        f"🙏 Подяки ({len(podyaka_out)})",
        redaktsiia_tab_label,
        f"⚠️ Помилки ({len(errors)})",
        "🔍 Лог",
    ])

    def _rows_to_display(rows, cols):
        return [
            {c: r.get(c, "") for c in cols} | {"⚠": "⚠" if r.get("warning") else ""}
            for r in rows
        ]

    with tab_z:
        st.caption("Зведена таблиця для типографії — усі документи разом")
        rows = []
        for i, r in enumerate(zvedena, 1):
            rows.append({
                "№": i,
                "№ Документу": r["num_doc"],
                "Тип": r["type"],
                "ПІБ": r["pib"],
                "К-сть": r["qty"],
                "ID угоди": r["id"],
                "": "⚠" if r.get("warning") else "",
            })
        st.dataframe(rows, use_container_width=True, hide_index=True,
                     column_config={
                         "№": st.column_config.NumberColumn(width="small"),
                         "К-сть": st.column_config.NumberColumn(width="small"),
                         "ID угоди": st.column_config.NumberColumn(width="medium"),
                         "": st.column_config.TextColumn(width="small"),
                     })

    with tab_d:
        st.caption("Таблиця дипломів для типографії")
        rows = []
        for i, r in enumerate(diploma_out, 1):
            rows.append({
                "№": i,
                "№ Диплому": r["num_doc"],
                "ПІБ учасника": r["pib"],
                "К-сть": r["qty"],
                "ID угоди": r["id"],
                "": "⚠" if r.get("warning") else "",
            })
        st.dataframe(rows, use_container_width=True, hide_index=True,
                     column_config={
                         "К-сть": st.column_config.NumberColumn(width="small"),
                         "ID угоди": st.column_config.NumberColumn(width="medium"),
                         "": st.column_config.TextColumn(width="small"),
                     })

    with tab_p:
        st.caption("Таблиця подяк для типографії")
        rows = []
        for i, r in enumerate(podyaka_out, 1):
            rows.append({
                "№": i,
                "№ Подяки": r["num_doc"],
                "ПІБ керівника": r["pib"],
                "К-сть": r["qty"],
                "ID угоди": r["id"],
                "": "⚠" if r.get("warning") else "",
            })
        st.dataframe(rows, use_container_width=True, hide_index=True,
                     column_config={
                         "К-сть": st.column_config.NumberColumn(width="small"),
                         "ID угоди": st.column_config.NumberColumn(width="medium"),
                         "": st.column_config.TextColumn(width="small"),
                     })

    with tab_r:
        if not redaktsiia:
            st.info("Завантажте файл редакції подяк у розділі «Редакція подяк (опційно)» і повторно запустіть агент.")
        else:
            in_druk_list  = [r for r in redaktsiia if r.get("in_druk")]
            not_druk_list = [r for r in redaktsiia if not r.get("in_druk")]
            c1, c2, c3 = st.columns(3)
            c1.metric("Всього у редакції", len(redaktsiia))
            c2.metric("✅ Є у списку друку", len(in_druk_list))
            c3.metric("— Не замовлено друк", len(not_druk_list))

            rows = []
            for r in redaktsiia:
                rows.append({
                    "№ ред.":          r["num_podyaka"],
                    "ПІБ керівника":   r["pib_kerivnyk"],
                    "School of arts":  r["school"],
                    "Учасники":        r["uchastnyky"],
                    "Друк":            "✅" if r.get("in_druk") else "—",
                    "№ подяки (друк)": r.get("druk_num") or "",
                    "Знайдений ПІБ":   r.get("druk_pib") or "",
                })
            st.dataframe(rows, use_container_width=True, hide_index=True,
                         column_config={
                             "№ ред.":          st.column_config.NumberColumn(width="small"),
                             "Друк":            st.column_config.TextColumn(width="small"),
                             "№ подяки (друк)": st.column_config.NumberColumn(width="medium"),
                             "ПІБ керівника":   st.column_config.TextColumn(width="large"),
                             "Знайдений ПІБ":   st.column_config.TextColumn(width="large"),
                         })

            # Кнопка скачати тільки ті, що є у друку
            if in_druk_list:
                import io as _io
                import openpyxl as _opxl
                from openpyxl.styles import PatternFill as _PF, Font as _Font
                wb_out = _opxl.Workbook()
                ws_out = wb_out.active
                ws_out.title = "Редакція для друку"
                hdr = ["№ ред.", "ПІБ керівника", "School of arts", "Учасники",
                       "№ подяки (друк)", "Знайдений ПІБ"]
                ws_out.append(hdr)
                for cell in ws_out[1]:
                    cell.font = _Font(bold=True)
                    cell.fill = _PF("solid", fgColor="4472C4")
                    cell.font = _Font(bold=True, color="FFFFFF")
                for r in in_druk_list:
                    ws_out.append([r["num_podyaka"], r["pib_kerivnyk"], r["school"], r["uchastnyky"],
                                   r.get("druk_num"), r.get("druk_pib")])
                for col, w in zip("ABCDEF", [10, 40, 45, 50, 18, 40]):
                    ws_out.column_dimensions[col].width = w
                buf = _io.BytesIO()
                wb_out.save(buf)
                st.download_button(
                    label=f"⬇ Завантажити редакцію для друку ({len(in_druk_list)} подяк)",
                    data=buf.getvalue(),
                    file_name=f"редакція_для_друку_{res['month'].replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    with tab_err:
        if errors:
            st.warning(f"{len(errors)} помилок. Перевірте вручну.")
            for e in errors:
                st.markdown(f"- {e}")
        else:
            st.success("Помилок не виявлено!")

    with tab_log:
        st.code(res["log"] or "(порожньо)", language=None)
