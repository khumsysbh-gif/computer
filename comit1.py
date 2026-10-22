import pandas as pd
import streamlit as st
#python -m streamlit run comit1.py
st.set_page_config(
    page_title="ແຜນການສອນວິຊາຄອມພິວເຕີ",
    page_icon="💻",
    layout="wide",
)

# ---------------------------------------------------------------
# ຂໍ້ມູນບົດຮຽນ
# ---------------------------------------------------------------
LESSONS = [
    {
        "title": "ບົດທີ 1: ຄວາມຮູ້ພື້ນຖານກ່ຽວກັບຄອມພິວເຕີ",
        "icon": "🖥️",
        "short": "ຄວາມຮູ້ພື້ນຖານຄອມພິວເຕີ",
        "time": "1 ຊົ່ວໂມງ",
        "objectives": [
            "ໃຫ້ນັກຮຽນຮູ້ຈັກຄອມພິວເຕີ ແລະ ປະໂຫຍດຂອງມັນ.",
            "ສາມາດບອກຊື່ສ່ວນປະກອບຫຼັກຂອງຄອມພິວເຕີ.",
            "ຮູ້ວິທີໃຊ້ຄອມພິວເຕີຢ່າງຖືກຕ້ອງ.",
        ],
        "content": [
            "ຄວາມໝາຍຂອງຄອມພິວເຕີ.",
            "ປະໂຫຍດຂອງຄອມພິວເຕີ.",
            "ສ່ວນປະກອບຫຼັກ: Monitor, CPU/System Unit, Keyboard ແລະ Mouse.",
            "ວິທີເປີດ ແລະ ປິດຄອມພິວເຕີຢ່າງປອດໄພ.",
        ],
        "activities": [
            "ຄູອະທິບາຍ ແລະ ສະແດງອຸປະກອນຕົວຈິງ.",
            "ນັກຮຽນຝຶກບອກຊື່ອຸປະກອນ.",
            "ນັກຮຽນຝຶກເປີດ ແລະ ປິດເຄື່ອງຕາມຂັ້ນຕອນ.",
        ],
        "assessment": [
            "ຖາມ-ຕອບກ່ຽວກັບສ່ວນປະກອບ.",
            "ໃຫ້ນັກຮຽນບອກປະໂຫຍດຂອງຄອມພິວເຕີ 3 ຢ່າງ.",
        ],
    },
    {
        "title": "ບົດທີ 2: ການຝຶກໃຊ້ Keyboard ແລະ Mouse",
        "icon": "⌨️",
        "short": "ການໃຊ້ Keyboard ແລະ Mouse",
        "time": "1 ຊົ່ວໂມງ",
        "objectives": [
            "ໃຫ້ນັກຮຽນຮູ້ຈັກປຸ່ມຕ່າງໆເທິງ Keyboard.",
            "ສາມາດໃຊ້ Mouse ເພື່ອຄວບຄຸມຕົວຊີ້ເທິງໜ້າຈໍ.",
            "ສາມາດພິມຕົວອັກສອນ ແລະ ຕົວເລກງ່າຍໆ.",
        ],
        "content": [
            "ປຸ່ມສຳຄັນ: Enter, Spacebar, Backspace, Shift ແລະ Caps Lock.",
            "ວິທີຈັບ ແລະ ເຄື່ອນຍ້າຍ Mouse.",
            "ການຄລິກຊ້າຍ, ຄລິກຂວາ ແລະ Double-click.",
            "ການພິມຕົວອັກສອນ ແລະ ຕົວເລກ.",
        ],
        "activities": [
            "ຄູສາທິດການໃຊ້ Keyboard ແລະ Mouse.",
            "ນັກຮຽນຝຶກຄລິກເປີດໂປຣແກຣມ.",
            "ນັກຮຽນຝຶກພິມຊື່ຂອງຕົນ ແລະ ຕົວເລກ 1–10.",
        ],
        "assessment": [
            "ໃຫ້ນັກຮຽນສາທິດການໃຊ້ Mouse.",
            "ໃຫ້ນັກຮຽນພິມປະໂຫຍກສັ້ນໆ 1 ປະໂຫຍກ.",
        ],
    },
    {
        "title": "ບົດທີ 3: ການໃຊ້ Microsoft Word ຂັ້ນພື້ນຖານ",
        "icon": "📝",
        "short": "Microsoft Word ຂັ້ນພື້ນຖານ",
        "time": "2 ຊົ່ວໂມງ",
        "objectives": [
            "ໃຫ້ນັກຮຽນຮູ້ຈັກໂປຣແກຣມ Microsoft Word.",
            "ສາມາດພິມ ແລະ ຈັດຮູບແບບຂໍ້ຄວາມໄດ້.",
            "ສາມາດບັນທຶກ ແລະ ເປີດເອກະສານຄືນໄດ້.",
        ],
        "content": [
            "ຄວາມໝາຍ ແລະ ປະໂຫຍດຂອງ Microsoft Word.",
            "ວິທີເປີດໂປຣແກຣມ ແລະ ສ້າງເອກະສານໃໝ່.",
            "ການພິມຂໍ້ຄວາມ, ປັບຂະໜາດຕົວອັກສອນ ແລະ ຈັດວາງຂໍ້ຄວາມ.",
            "ການບັນທຶກເອກະສານດ້ວຍ Save ແລະ Save As.",
        ],
        "activities": [
            "ຄູສາທິດການເປີດ Microsoft Word.",
            "ນັກຮຽນພິມຫົວຂໍ້ “ປະຫວັດຂອງຂ້ອຍ”.",
            "ນັກຮຽນພິມຂໍ້ຄວາມ 3–5 ແຖວ ແລະ ປັບຂະໜາດຕົວອັກສອນ.",
            "ນັກຮຽນບັນທຶກໄຟລ໌ໃສ່ໂຟນເດີທີ່ຄູກຳນົດ.",
        ],
        "assessment": [
            "ກວດເບິ່ງເອກະສານທີ່ນັກຮຽນສ້າງ.",
            "ກວດວ່ານັກຮຽນສາມາດບັນທຶກ ແລະ ເປີດໄຟລ໌ຄືນໄດ້ຫຼືບໍ່.",
        ],
    },
]

SECTIONS = [
    ("🎯 ຈຸດປະສົງ", "objectives"),
    ("📚 ເນື້ອໃນການສອນ", "content"),
    ("🧑‍🏫 ກິດຈະກຳການຮຽນ", "activities"),
    ("✅ ການປະເມີນຜົນ", "assessment"),
]


def render_lesson(lesson: dict, key_prefix: str) -> None:
    """ສະແດງບົດຮຽນ 1 ບົດ"""
    st.subheader(f"{lesson['icon']} {lesson['title']}")
    st.caption(f"⏱️ ເວລາແນະນຳ: {lesson['time']}")

    for label, field in SECTIONS:
        with st.expander(label, expanded=True):
            for i, item in enumerate(lesson[field], start=1):
                st.markdown(f"{i}. {item}")


def summary_table() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "ບົດຮຽນ": [f"ບົດທີ {i}" for i in range(1, len(LESSONS) + 1)],
            "ເນື້ອໃນ": [l["short"] for l in LESSONS],
            "ເວລາແນະນຳ": [l["time"] for l in LESSONS],
        }
    )


# ---------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------
st.sidebar.title("💻 ເມນູ")
view = st.sidebar.radio(
    "ເລືອກການສະແດງ",
    ["ທຸກບົດ (Tabs)", "ເລືອກບົດດຽວ", "ຕາຕະລາງສະຫຼຸບ"],
)

st.sidebar.divider()
st.sidebar.markdown("**ລະດັບນັກຮຽນ (ເລືອກເພື່ອບັນທຶກ)**")
level = st.sidebar.selectbox(
    "ນຳໄປສອນນັກຮຽນລະດັບໃດ?",
    ["ປະຖົມ", "ມັດທະຍົມ", "ວິຊາຊີບ / ວິທະຍາໄລ", "ຜູ້ເລີ່ມຕົ້ນທົ່ວໄປ"],
)

# ---------------------------------------------------------------
# ໜ້າຫຼັກ
# ---------------------------------------------------------------
st.title("📘 ແຜນການສອນວິຊາຄອມພິວເຕີ 3 ບົດ")
st.write(
    "ບົດສອນນີ້ມີຈຸດປະສົງ, ເນື້ອໃນການສອນ, ກິດຈະກຳການຮຽນ ແລະ "
    "ການປະເມີນຜົນ ສາມາດນຳໄປປັບໃຊ້ໃນຫ້ອງຮຽນໄດ້."
)
st.info(f"ກຸ່ມເປົ້າໝາຍ: **{level}**")

total_hours = sum(int(l["time"].split()[0]) for l in LESSONS)
c1, c2, c3 = st.columns(3)
c1.metric("ຈຳນວນບົດ", len(LESSONS))
c2.metric("ເວລາລວມ", f"{total_hours} ຊົ່ວໂມງ")
c3.metric("ກຸ່ມເປົ້າໝາຍ", level)

st.divider()

if view == "ທຸກບົດ (Tabs)":
    tabs = st.tabs([f"{l['icon']} ບົດທີ {i}" for i, l in enumerate(LESSONS, 1)])
    for i, (tab, lesson) in enumerate(zip(tabs, LESSONS)):
        with tab:
            render_lesson(lesson, f"tab{i}")

elif view == "ເລືອກບົດດຽວ":
    choice = st.selectbox("ເລືອກບົດຮຽນ", [l["title"] for l in LESSONS])
    lesson = next(l for l in LESSONS if l["title"] == choice)
    render_lesson(lesson, "single")

else:
    st.subheader("📊 ຕາຕະລາງສະຫຼຸບ 3 ບົດ")
    st.dataframe(summary_table(), hide_index=True, use_container_width=True)

st.divider()
st.caption("ສ້າງດ້ວຍ Streamlit • ແຜນການສອນວິຊາຄອມພິວເຕີ")