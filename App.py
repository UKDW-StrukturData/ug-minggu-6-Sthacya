import streamlit as st
from DatabaseManager import excelManager


em = excelManager("dataExcel.xlsx")
options = ["Choose Action","Insert", "Edit", "Delete"]
choice = st.selectbox("Choose an action:", options)
saveChange = st.checkbox("SaveChanges",value=False)

if choice in ("Edit", "Delete"):
    nim = st.text_input("Enter targeted NIM:", key="targetNim")
    if (choice == "Delete"):
        if (sum([1 for i in nim if str(i).isalpha()]) > 0): st.error("Input nim harus angka semua")
        elif st.button("Delete"):
            if (not em.getData("NIM",nim)): st.error("nim not found")
            else:
                em.deleteData(nim,saveChange)
                if (not em.getData("NIM",nim)): st.success("deleted")
            

if (choice in ("Insert","Edit")):
    newNim = st.text_input("Enter New NIM:",key="newNim")
    newName = st.text_input("Enter New Name:",key="newName")
    newGrade = st.text_input("Enter New Grade :", key="newGrade")

    if (choice == "Edit"):
        if st.button("Edit"):
            if (sum([1 for i in newNim if str(i).isalpha()]) > 0): st.error("Input nim harus angka semua")
            elif (sum([1 for i in newName if str(i).isdigit()]) > 0): st.error("Input nama harus Alphabet semua")
            elif (sum([1 for i in newGrade if str(i).isalpha()]) > 0): st.error("Input nilai harus angka semua")
            else:
                if (not em.getData("NIM",nim)): st.success("Nim not found")
                else:
                    result = em.editData(str(nim),{"NIM":str(newNim).strip(),"Nama":str(newName).strip(),"Nilai":int(newGrade.strip())},saveChange)
                    if (em.getData("NIM",newNim)): st.success("edited")
                    else: st.error("edit failed")

    if (choice == "Insert"):    
        if (st.button("Insert")):
            if (sum([1 for i in newNim if str(i).isalpha()]) > 0): st.error("Input nim harus angka semua")
            elif (sum([1 for i in newName if str(i).isdigit()]) > 0): st.error("Input nama harus Alphabet semua")
            elif (sum([1 for i in newGrade if str(i).isalpha()]) > 0): st.error("Input nilai harus angka semua")
            else:
                if (em.getData("NIM",newNim)): st.error("Nim already exist")
                else:
                    em.insertData({"NIM":str(newNim).strip(),"Nama":str(newName).strip(),"Nilai":int(newGrade.strip())},saveChange)
                    if (em.getData("NIM",newNim)): st.success("inserted")
                    else: st.error("insert fail")
                                
# TODO: buatkan sistem filter data tabel berdasarkan kolom yang memiliki data angka
option = ["None",">","<","=","<=",">="]
filterSelectBox = st.selectbox("Opsi Filter: ",option)

if filterSelectBox == "None":
    st.table(em.getDataFrame())  # tabel biasa
else:
    targetFilterColumn = st.selectbox("Target Column", ["NIM", "Nilai"])  # pilihan kolom yang bisa difilter
    filter_value = st.text_input("Filter Nilai") 

    if filter_value != "":
        try:
            filter_num = int(filter_value)
            df = em.getDataFrame()

            if filterSelectBox == ">":
                filtered_df = df[df[targetFilterColumn] > filter_num]
            elif filterSelectBox == "<":
                filtered_df = df[df[targetFilterColumn] < filter_num]
            elif filterSelectBox == "=":
                filtered_df = df[df[targetFilterColumn] == filter_num]
            elif filterSelectBox == "<=":
                filtered_df = df[df[targetFilterColumn] <= filter_num]
            elif filterSelectBox == ">=":
                filtered_df = df[df[targetFilterColumn] >= filter_num]

            st.table(filtered_df)
        except ValueError:
            st.error("Filter harus berupa angka")
    else:
        st.info("Masukkan nilai filter")

if (filter != ""):
    if (filterSelectBox == ">"):
        st.table(em.getDataFrame()[em.getDataFrame()[targetFilterColumn] > int(filter)]) 
    # TODO: lanjutkan code di atas
    # note: cara filter ada di modul
    try:
        filter_value = int(filter)
        df = em.getDataFrame()
        
        if filterSelectBox == ">":
            st.table(df[df[targetFilterColumn] > filter_value])
        elif filterSelectBox == "<":
            st.table(df[df[targetFilterColumn] < filter_value])
        elif filterSelectBox == "=":
            st.table(df[df[targetFilterColumn] == filter_value])
        elif filterSelectBox == "<=":
            st.table(df[df[targetFilterColumn] <= filter_value])
        elif filterSelectBox == ">=":
            st.table(df[df[targetFilterColumn] >= filter_value])
    except ValueError:
        st.error("Filter harus berupa angka")
else:
    st.info("Masukkan nilai filter")
