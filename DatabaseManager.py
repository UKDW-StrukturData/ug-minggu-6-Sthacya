import pandas

class excelManager:
    def __init__(self,filePath:str,sheetName:str="Sheet1"):
        self.__filePath = filePath
        self.__sheetName = sheetName
        self.__data = pandas.read_excel(filePath,sheet_name=sheetName)


    def insertData(self,newData:dict,saveChange:bool=False):
        # Validasi NIM tidak boleh duplikat
        nim_baru = newData.get("NIM", "").strip()
        nama_baru = newData.get("Nama", "").strip()

        # Cek NIM sudah ada atau belum
        if not nim_baru:
            print("NIM tidak boleh kosong")
            return

        # Cek Nama harus bukan angka (ada huruf)
        if any(char.isdigit() for char in nama_baru):
            print("Nama tidak boleh angka")
            return
        
        # Cek duplikat NIM
        if nim_baru in self.__data["NIM"].astype(str).values:
            print("Nim Sudah Ada")
            return
        
        # Insert data baru
        df_new_row = pandas.DataFrame([newData])
        self.__data = pandas.concat([self.__data, df_new_row], ignore_index=True)

        if saveChange:
            self.saveChange()
        
        print("Data Sukses di Masukan")
    
    def deleteData(self, targetedNim:str,saveChange:bool=False):
        targetedNim = targetedNim.strip()

        # Cek apakah NIM ada di data
        if targetedNim not in self.__data["NIM"].astype(str).values:
            print("Nim tidak ditemukan")
            return
        
        # Dapatkan index baris yang punya nim tersebut
        idx = self.__data.index[self.__data["NIM"].astype(str) == targetedNim].tolist()
        
        # Drop baris tersebut
        self.__data.drop(idx, inplace=True)
        self.__data.reset_index(drop=True, inplace=True)

        if saveChange:
            self.saveChange()

        print("Data Sukses di Hapus")
    
    def editData(self, targetedNim:str, newData:dict,saveChange:bool=False) -> dict:
        targetedNim = targetedNim.strip()
        new_nim = newData.get("NIM", "").strip()
        new_nama = newData.get("Nama", "").strip()

        # Cek apakah NIM yang akan diedit ada
        if targetedNim not in self.__data["NIM"].astype(str).values:
            print("Nim tidak ditemukan")
            return None
        
        # Validasi nama baru tidak boleh angka
        if any(char.isdigit() for char in new_nama):
            print("Nama tidak boleh angka")
            return None
        
        # Dapatkan index baris yang akan di edit
        idx = self.__data.index[self.__data["NIM"].astype(str) == targetedNim].tolist()[0]

        # Jika new_nim berbeda dan sudah ada di data lain, tolak edit
        if new_nim != targetedNim and new_nim in self.__data["NIM"].astype(str).values:
            print("Nim Sudah Ada")
            return None
        
        # Lakukan edit
        self.__data.at[idx, "NIM"] = new_nim
        self.__data.at[idx, "Nama"] = new_nama

        if saveChange:
            self.saveChange()

        print("Data Sukses di Edit")
        return self.__data.loc[idx].to_dict()
    
                    
    def getData(self, colName:str, data:str) -> dict:
        collumn = self.__data.columns # mendapatkan list dari nama kolom tabel
        
        # cari index dari nama kolom dan menjaganya dari typo atau spasi berlebih
        collumnIndex = [i for i in range(len(collumn)) if (collumn[i].lower().strip() == colName.lower().strip())] 
        
        # validasi jika input kolom tidak ada pada data excel
        if (len(collumnIndex) != 1): return None
        
        # nama kolom yang sudah pasti benar dan ada
        colName = collumn[collumnIndex[0]]
        
        
        resultDict = dict() # tempat untuk hasil
        
        for i in self.__data.index: # perulangan ke baris tabel
            cellData = str(self.__data.at[i,colName]) # isi tabel yand dijadikan str
            if (cellData == data): # jika data cell sama dengan data input
                for col in collumn: # perulangan ke nama-nama kolom
                    resultDict.update({str(col):str(self.__data.at[i,col])}) # masukan data {namaKolom : data pada cell} ke resultDict
                resultDict.update({"Row":i}) # tambahkan row nya pada resultDict
                return resultDict # kembalikan resultDict
        
        return None
    
    def saveChange(self):
        self.__data.to_excel(self.__filePath, sheet_name=self.__sheetName , index=False)
    
    def getDataFrame(self):
        return self.__data
