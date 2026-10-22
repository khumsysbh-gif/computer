import os
import streamlit as st
#python -m streamlit run comit3.py
st.set_page_config(page_title="ບົດຮຽນຄອມພິວເຕີເບື້ອງຕົ້ນ", page_icon="💻", layout="centered")

# ---------------------------------------------------------------
# ຂໍ້ມູນບົດຮຽນ
# ໃສ່ຮູບໄວ້ໃນໂຟນເດີ images/ ຊື່ img1.png, img2.png, img3.png
# ຖ້າບໍ່ມີໄຟລ໌ ໂປຣແກຣມຈະສະແດງຄຳແນະນຳ ແລະ ໃຫ້ອັບໂຫຼດຮູບໄດ້ເອງ
# ---------------------------------------------------------------
LESSONS = {
    "ບົດທີ 1: ຄວາມຮູ້ພື້ນຖານກ່ຽວກັບຄອມພິວເຕີ": {
        "image": "images/img1.png",
        "image_tip": "ຮູບພາບລວມຂອງຊຸດຄອມພິວເຕີຕັ້ງໂຕະ (Desktop) ທີ່ປະກອບດ້ວຍໜ້າຈໍ, ເຄສ, ແປ້ນພິມ ແລະ ເມົ້າ ພ້ອມລູກສອນຊີ້ບອກແຕ່ລະສ່ວນ.",
    },
    "ບົດທີ 2: ການໃຊ້ງານແປ້ນພິມ ແລະ ເມົ້າ": {
        "image": "images/img2.png",
        "image_tip": "ຮູບຊູມຂະຫຍາຍປຸ່ມເທິງແປ້ນພິມ (ເນັ້ນ Enter, Spacebar, Backspace) ພ້ອມຮູບເມົ້າ ແລະ ການວາງມື.",
    },
    "ບົດທີ 3: ການໃຊ້ງານ Microsoft Word ເບື້ອງຕົ້ນ": {
        "image": "images/img3.png",
        "image_tip": "Screenshot ຂອງ Microsoft Word ເນັ້ນແຖບ Home (Ribbon) ບ່ອນເລືອກ Font ແລະ ຂະໜາດ.",
    },
}


def show_image(path: str, tip: str, key: str):
    """ສະແດງຮູບປະກອບ ຫຼື ຈຸດແນະນຳໃຫ້ໃສ່ຮູບ."""
    if os.path.exists(path):
        st.image(path, use_container_width=True)
        return
    st.info(f"🖼️ **ຄຳແນະນຳໃສ່ຮູບປະກອບ:** {tip}")
    uploaded = st.file_uploader("ອັບໂຫຼດຮູບປະກອບ", type=["png", "jpg", "jpeg"], key=key)
    if uploaded:
        st.image(uploaded, use_container_width=True)


def lesson_1():
    st.header("ບົດທີ 1: ຄວາມຮູ້ພື້ນຖານກ່ຽວກັບຄອມພິວເຕີ (Computer Basics)")
    st.subheader("1. ຄວາມໝາຍຂອງຄອມພິວເຕີ")
    st.write(
        "ຄອມພິວເຕີ (Computer) ຄື ອຸປະກອນອີເລັກໂຕຣນິກອັດຕະໂນມັດ ທີ່ຖືກສ້າງຂຶ້ນມາເພື່ອປະຕິບັດງານຕາມຄຳສັ່ງ, "
        "ຮັບຂໍ້ມູນ (Input) ເຂົ້າໄປ, ນຳມາປະມວນຜົນ (Process) ແລະ ສະແດງຜົນອອກມາ (Output) "
        "ໃຫ້ມະນຸດສາມາດນຳໄປໃຊ້ປະໂຫຍດໄດ້."
    )
    st.subheader("2. ອົງປະກອບຫຼັກຂອງຄອມພິວເຕີ")
    st.write("ຄອມພິວເຕີແບ່ງອອກເປັນ 2 ພາກສ່ວນຫຼັກຄື:")
    st.markdown(
        "- **ຮາດແວ (Hardware):** ພາກສ່ວນໂຄງສ້າງທາງກາຍະພາບທີ່ຈັບຕ້ອງໄດ້ ເຊັ່ນ: ໜ້າຈໍ, ແປ້ນພິມ, ເມົ້າ ແລະ ເຄສ (Case)\n"
        "- **ຊອບແວ (Software):** ຊຸດຄຳສັ່ງ ຫຼື ໂປຣແກຣມທີ່ສັ່ງໃຫ້ຮາດແວເຮັດວຽກ ເຊັ່ນ: ລະບົບປະຕິບັດການ Windows, ໂປຣແກຣມ Microsoft Word"
    )


def lesson_2():
    st.header("ບົດທີ 2: ການໃຊ້ງານແປ້ນພິມ ແລະ ເມົ້າ (Keyboard & Mouse)")
    st.subheader("1. ການໃຊ້ງານເມົ້າ (Mouse)")
    st.write("ເມົ້າແມ່ນອຸປະກອນທີ່ໃຊ້ສຳລັບຊີ້ນຳ ແລະ ຄວບຄຸມການເຮັດວຽກເທິງໜ້າຈໍ, ປະກອບດ້ວຍປຸ່ມຫຼັກ:")
    st.markdown(
        "- **ຄລິກຊ້າຍ (Left Click):** ເລືອກ (Select) ໄອຄອນ ຫຼື ເປີດໂປຣແກຣມ (ດັບເບິນຄລິກ)\n"
        "- **ຄລິກຂວາ (Right Click):** ເປີດເມນູລັດ (Context Menu) ເພື່ອເບິ່ງຄຳສັ່ງເພີ່ມເຕີມ\n"
        "- **ລໍ້ເລື່ອນ (Scroll Wheel):** ເລື່ອນໜ້າເວັບ ຫຼື ເອກະສານຂຶ້ນ-ລົງ"
    )
    st.subheader("2. ການໃຊ້ງານແປ້ນພິມ (Keyboard)")
    st.write("ແປ້ນພິມໃຊ້ສຳລັບພິມຂໍ້ຄວາມ ແລະ ປ້ອນຄຳສັ່ງຕ່າງໆ, ໂດຍມີກຸ່ມປຸ່ມສຳຄັນ:")
    st.markdown(
        "- **Spacebar:** ປຸ່ມຍາວສຸດ ໃຊ້ສຳລັບເຄື່ອນທີ່ວ່າງ\n"
        "- **Enter:** ລົງແຖວໃໝ່ ຫຼື ຢືນຢັນຄຳສັ່ງ\n"
        "- **Backspace / Delete:** ລຶບຕົວອັກສອນ\n"
        "- **Shift / Ctrl / Alt:** ປຸ່ມພິເສດທີ່ກົດຮ່ວມກັບປຸ່ມອື່ນເພື່ອເປັນທາງລັດ (Shortcut)"
    )


def lesson_3():
    st.header("ບົດທີ 3: ການໃຊ້ງານໂປຣແກຣມພິມເອກະສານ (Microsoft Word ເບື້ອງຕົ້ນ)")
    st.subheader("1. ການເປີດ ແລະ ສ້າງເອກະສານໃໝ່")
    st.markdown(
        "1. ຄລິກທີ່ປຸ່ມ **Start** ຢູ່ມຸມຊ້າຍລຸ່ມຂອງໜ້າຈໍ\n"
        "2. ຊອກຫາ ແລະ ຄລິກເລືອກໂປຣແກຣມ **Word**\n"
        "3. ເລືອກ **Blank Document** ເພື່ອເລີ່ມຕົ້ນສ້າງເອກະສານເປົ່າ"
    )
    st.subheader("2. ການຈັດຮູບແບບຕົວອັກສອນພື້ນຖານ")
    st.write("ຫຼັງຈາກພິມຂໍ້ຄວາມແລ້ວ, ສາມາດປ່ຽນຮູບແບບຕົວໜັງສືໄດ້ທີ່ແຖບ **Home**:")
    st.markdown(
        "- **Font:** ເລືອກແບບຕົວອັກສອນ (ຕົວຢ່າງ: Phetsarath OT)\n"
        "- **Font Size:** ປັບຂະໜາດຕົວໜັງສືໃຫ້ໃຫຍ່ ຫຼື ນ້ອຍລົງ\n"
        "- **B (Bold), I (Italic), U (Underline):** ຕົວໜາ, ຕົວອຽງ ແລະ ຂີດເສັ້ນໃຕ້"
    )


RENDERERS = [lesson_1, lesson_2, lesson_3]

# ---------------------------------------------------------------
# ສ່ວນສະແດງຜົນຫຼັກ
# ---------------------------------------------------------------
st.title("💻 ບົດຮຽນວິຊາຄອມພິວເຕີເບື້ອງຕົ້ນ")

with st.sidebar:
    st.header("ສາລະບານ")
    choice = st.radio("ເລືອກບົດຮຽນ", list(LESSONS.keys()))

idx = list(LESSONS.keys()).index(choice)
RENDERERS[idx]()

info = LESSONS[choice]
show_image(info["image"], info["image_tip"], key=f"upload_{idx}")

st.divider()
col1, col2 = st.columns(2)
with col1:
    if idx > 0:
        st.caption(f"⬅️ ກ່ອນໜ້າ: {list(LESSONS.keys())[idx - 1]}")
with col2:
    if idx < len(LESSONS) - 1:
        st.caption(f"ຖັດໄປ: {list(LESSONS.keys())[idx + 1]} ➡️")