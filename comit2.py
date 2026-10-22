import os
import streamlit as st
#python -m streamlit run comit2.py
st.set_page_config(page_title="ອຸປະກອນຂອງຄອມພິວເຕີ", page_icon="💻", layout="wide")

IMG_DIR = "images"

# ແຕ່ລະຫົວຂໍ້: ຊື່, ໄຟລ໌ຮູບ, ຄຳອະທິບາຍ, ຈຸດສຳຄັນ
TOPICS = [
    {
        "title": "1. ຄອມພິວເຕີ (Computer)",
        "image": "computer.jpg",
        "caption": "ຄອມພິວເຕີຕັ້ງໂຕະ (Desktop PC) ພ້ອມຈໍສະແດງຜົນ, ຄີບອດ ແລະ ເມົ້າ",
        "desc": (
            "ຄອມພິວເຕີແມ່ນເຄື່ອງເອເລັກໂຕຣນິກທີ່ຮັບຂໍ້ມູນ (Input), ປະມວນຜົນ (Process), "
            "ເກັບຮັກສາ (Storage) ແລະ ສະແດງຜົນ (Output) ຕາມຄຳສັ່ງຂອງຜູ້ໃຊ້."
        ),
        "points": [
            "ອຸປະກອນຮັບຂໍ້ມູນ: ຄີບອດ, ເມົ້າ",
            "ອຸປະກອນສະແດງຜົນ: ຈໍພາບ (Monitor)",
            "ຕົວເຄື່ອງ (Case) ເປັນບ່ອນໃສ່ຊິ້ນສ່ວນພາຍໃນທັງໝົດ",
        ],
    },
    {
        "title": "2. ໂປຣເຊັດເຊີ (CPU)",
        "image": "cpu.jpg",
        "caption": "CPU ຕົວຢ່າງ: Intel Core i9-14900KF",
        "desc": (
            "CPU (Central Processing Unit) ແມ່ນ 'ສະໝອງ' ຂອງຄອມພິວເຕີ "
            "ເຮັດໜ້າທີ່ປະມວນຜົນຄຳສັ່ງ ແລະ ຄຳນວນຕ່າງໆ."
        ),
        "points": [
            "ຄວາມໄວວັດເປັນ GHz",
            "ມີຫຼາຍແກນ (Core) ແລະ ເທຣດ (Thread) ເພື່ອເຮັດວຽກພ້ອມກັນ",
            "ຜູ້ຜະລິດຫຼັກ: Intel ແລະ AMD",
        ],
    },
    {
        "title": "3. ກາດຈໍ (GPU)",
        "image": "gpu.jpg",
        "caption": "ກາດຈໍ (Graphics Card) ຍີ່ຫໍ້ ASUS ROG Strix ແບບພັດລົມຄູ່",
        "desc": (
            "GPU (Graphics Processing Unit) ເຮັດໜ້າທີ່ປະມວນຜົນພາບ ແລະ ກຣາຟິກ "
            "ເໝາະກັບການຫຼິ້ນເກມ, ຕັດຕໍ່ວິດີໂອ, ງານ 3D ແລະ AI."
        ),
        "points": [
            "ມີໜ່ວຍຄວາມຈຳຂອງຕົນເອງ (VRAM)",
            "ມີພັດລົມລະບາຍຄວາມຮ້ອນ",
            "ມີພອດຕໍ່ຈໍ ເຊັ່ນ HDMI ແລະ DisplayPort",
        ],
    },
    {
        "title": "4. ເມນບອດ (Motherboard)",
        "image": "motherboard.jpg",
        "caption": "ເມນບອດ (Mainboard) ທີ່ມີຊ່ອງໃສ່ CPU, RAM ແລະ ກາດຂະຫຍາຍ",
        "desc": (
            "ເມນບອດແມ່ນແຜງວົງຈອນຫຼັກທີ່ເຊື່ອມຕໍ່ອຸປະກອນທຸກຢ່າງເຂົ້າດ້ວຍກັນ "
            "ແລະ ສົ່ງໄຟຟ້າ/ຂໍ້ມູນໄປຫາແຕ່ລະຊິ້ນສ່ວນ."
        ),
        "points": [
            "ມີຊ່ອງ CPU Socket",
            "ມີຊ່ອງໃສ່ RAM ແລະ PCI Express",
            "ມີພອດຕ່າງໆ ເຊັ່ນ USB, LAN, Audio",
        ],
    },
    {
        "title": "5. ແຣມ (RAM)",
        "image": "ram.jpg",
        "caption": "ໜ່ວຍຄວາມຈຳ RAM ແບບ 2 ແຜງ (Dual Channel)",
        "desc": (
            "RAM (Random Access Memory) ແມ່ນໜ່ວຍຄວາມຈຳຊົ່ວຄາວ "
            "ທີ່ເກັບຂໍ້ມູນຂອງໂປຣແກຣມທີ່ກຳລັງໃຊ້ງານ ເພື່ອໃຫ້ CPU ເຂົ້າເຖິງໄດ້ໄວ."
        ),
        "points": [
            "ຂໍ້ມູນຈະຫາຍເມື່ອປິດເຄື່ອງ (Volatile)",
            "ຄວາມຈຸຍິ່ງຫຼາຍ ຍິ່ງເປີດໂປຣແກຣມພ້ອມກັນໄດ້ຫຼາຍ",
            "ປະເພດທີ່ນິຍົມ: DDR4 ແລະ DDR5",
        ],
    },
    {
        "title": "6. ຄລາວ ແລະ ເຊີບເວີ (Cloud / Server)",
        "image": "cloud.jpg",
        "caption": "ຫ້ອງເຊີບເວີ (Data Center) ແລະ ສັນຍາລັກ Cloud",
        "desc": (
            "Cloud ແມ່ນບໍລິການເກັບຂໍ້ມູນ ແລະ ປະມວນຜົນຜ່ານອິນເຕີເນັດ "
            "ໂດຍໃຊ້ເຊີບເວີຈຳນວນຫຼາຍໃນ Data Center."
        ),
        "points": [
            "ອັບໂຫຼດ/ດາວໂຫຼດຂໍ້ມູນໄດ້ທຸກບ່ອນ",
            "ສຳຮອງຂໍ້ມູນ (Backup) ປອດໄພ",
            "ຕົວຢ່າງ: Google Drive, OneDrive, AWS",
        ],
    },
]


def show_image(filename: str, caption: str):
    path = os.path.join(IMG_DIR, filename)
    if os.path.exists(path):
        st.image(path, caption=caption, use_container_width=True)
    else:
        st.warning(f"ບໍ່ພົບຮູບ: {path}  — ກະລຸນາວາງໄຟລ໌ຮູບໃນໂຟນເດີ '{IMG_DIR}'")


def show_topic(t: dict):
    st.header(t["title"])
    col1, col2 = st.columns([1, 1])
    with col1:
        show_image(t["image"], t["caption"])
    with col2:
        st.subheader("ຄຳອະທິບາຍ")
        st.write(t["desc"])
        st.subheader("ຈຸດສຳຄັນ")
        for p in t["points"]:
            st.markdown(f"- {p}")


# ---------- ສ່ວນສະແດງຜົນ ----------
st.title("💻 ອຸປະກອນຂອງຄອມພິວເຕີ")
st.caption("ຮຽນຮູ້ 6 ຫົວຂໍ້ກ່ຽວກັບອຸປະກອນຄອມພິວເຕີ")

menu = st.sidebar.radio(
    "ເລືອກຫົວຂໍ້",
    options=range(len(TOPICS)),
    format_func=lambda i: TOPICS[i]["title"],
)
st.sidebar.markdown("---")
st.sidebar.info("ວາງຮູບໄວ້ໃນໂຟນເດີ `images/`")

show_topic(TOPICS[menu])

with st.expander("📋 ເບິ່ງທຸກຫົວຂໍ້ແບບຫຍໍ້"):
    cols = st.columns(3)
    for i, t in enumerate(TOPICS):
        with cols[i % 3]:
            show_image(t["image"], t["title"])