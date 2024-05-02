# copyrigth by Dwinanda AS
# Author: Dwinanda Alfauzan Suhando
# Tahun: 2022

from array import array
from ast import And
from codecs import ignore_errors
from csv import writer
from ctypes import sizeof
from itertools import count
from re import T
import string
from turtle import pen
import pandas as pd
import math
import timeit

#====== Variabel Global ===========
kumpulan_kurs = []
kurs_array = []
array_kurs_nilai_FLRG = []
array_hasil_peramalan = []
kumpulan_nilai_tengah = []
nilai_tengah_interval = []
kumpulan_kelas = []
kumpulan_fuzzyfikasi = []
FLRData_kumpulan = []
perpus_matrix = []
array_nilai_matrix = []
array_of_matrix = []
perpus = []
fuzzyfikasi = []
fuzzyfikasi_data_historis = []

###### MARKOV CHAIN ######

#Read Excel
kurs = pd.read_excel(r'D:\Kuliah\Semester8\TA\TA_KoTA305\TA\DataHistorisKursJualKetika.xlsx', sheet_name='Kurs Jual Ketika Covid')
data = pd.DataFrame(kurs, columns = ['No', 'Kurs'])
datav2= pd.DataFrame(kurs, columns = ['No', 'Tanggal', 'Kurs', 'Fuzzyfikasi'])
# print(kurs_array)
# pd.set_option("display.max_rows", None)

print(data)
row_count = pd.DataFrame(kurs, columns = ['No']).count()
# print(row_count[0])


def intervalKelas():
    #Tahapan menentukan Nilai Minimum dan Maksimum
    minimal=pd.DataFrame(kurs, columns = ['Kurs']).min().astype(int)
    d1 = minimal[0] - 0 # Nilai berdasarkan dari kombinasi minimum, maksimum, dan rata-rata dari delta kurs dollar
    maksimal=pd.DataFrame(kurs, columns = ['Kurs']).max().astype(int)
    d2 = maksimal[0] + 0 # Nilai berdasarkan dari kombinasi minimum, maksimum, dan rata-rata dari delta kurs dollar
    print('Nilai Minimal dari data historis yaitu: ', d1)
    print('Nilai Maksimal dari data historis yaitu: ', d2)
    #Tahapan menentukan Banyak Kelas, Rentang Kelas, dan Interval Kelas
    banyak_kelas = 1+(3.3*math.log10(row_count))
    banyak_kelas = int(banyak_kelas)
    print('Banyak Kelas adalah:', banyak_kelas)
    rentang_kelas = d2 - d1 
    print('Rentang Kelas dari data historis yaitu:', rentang_kelas)
    interval_kelas = rentang_kelas/banyak_kelas
    interval_kelas = interval_kelas.astype(int)
    print('Interval Kelas dari data historis yaitu:', interval_kelas)
    #Pembentukan Interval Kelas data historis
    for i in range(banyak_kelas):
        k = 1
        i = d1
    for j in range(banyak_kelas):
        kelas = []
        j = i + interval_kelas
        kelas.append(i)
        kelas.append(j)
        kelas.append(f"A{k}")
        kumpulan_fuzzyfikasi.append(f"A{k}")
        kumpulan_kelas.append(kelas)
        print(kelas)
        nilai = (i + j)/2
        nilai_tengah_interval.append(nilai)
        kelas_dan_nilai_tengah = []
        kelas_dan_nilai_tengah.append(f"A{k}")
        kelas_dan_nilai_tengah.append(nilai)
        kumpulan_nilai_tengah.append(kelas_dan_nilai_tengah)
        print(f"Interval Kelas ke - {k}: ", i, ' - ', j, " = " f"A{k}", " = ", nilai)
        # j = j + 1
        i = j
        k = k + 1
print(nilai_tengah_interval)
print(kumpulan_nilai_tengah)

# print(kumpulan_kelas)
print(kumpulan_fuzzyfikasi)

def pengelompokkanFuzzyfikasiKurs():
    data1 = pd.DataFrame(kurs, columns = ['Kurs'])
    print(data1)
    # print(data1.values[0][0])

    # Loop Penentuan Fuzzyfikasi
    for kurs_b4_covid in kurs.values:
        kurs_checked = kurs_b4_covid[2]
        for kelas in kumpulan_kelas:
            if kurs_checked <= kelas[1]:
                data_historis = []
                data_historis.append(kurs_b4_covid[0])
                data_historis.append(kurs_b4_covid[1])
                data_historis.append(kurs_b4_covid[2])
                data_historis.append(kelas[2])
                fuzzyfikasi.append(data_historis)
                # print(str(kurs_checked) + " masuk ke kelas " + kelas[2])
                fuzzyfikasi_kurs_data_historis = []
                fuzzyfikasi_kurs_data_historis.append(kurs_checked)
                fuzzyfikasi_kurs_data_historis.append(kelas[2])
                fuzzyfikasi_data_historis.append(fuzzyfikasi_kurs_data_historis)     
                break
    Fuzzyfikasi = pd.DataFrame(fuzzyfikasi, columns=['No', 'Tanggal', 'Kurs', 'Fuzzyfikasi'])
    with pd.ExcelWriter('D:\Kuliah\Semester8\TA\TA_KoTA305\TA\DataHistorisKursJualKetika.xlsx', mode='a') as writer:
        Fuzzyfikasi.to_excel(writer, sheet_name='Fuzzyfikasi Kurs 0000', index=False)
    kurs_array.append(data1)
    kumpulan_kurs.append(kurs_array)

def hitungFLR():
    # Penentuan FLR
    FLR = pd.read_excel(r'D:\Kuliah\Semester8\TA\TA_KoTA305\TA\DataHistorisKursJualKEtika.xlsx', sheet_name='Fuzzyfikasi Kurs 0000')
    FLRData = pd.DataFrame(FLR, columns=['Fuzzyfikasi'])
    for i in range(0, FLRData.size - 1):
        if i+1 <= FLRData.size - 1:
            FLRData_values_kumpulan = []
            if FLRData.values[1][0] == FLRData.values[1][0]:
                # FLRData_values_kumpulan.append(FLRData.values[i][0]), " - ", FLRData.values[i+1][0])
                FLRData_values_kumpulan.append(FLRData.values[i][0])
                FLRData_values_kumpulan.append(FLRData.values[i+1][0])
                FLRData_kumpulan.append(FLRData_values_kumpulan)
    #print(kumpulan_fuzzyfikasi)

def hitungFLRG():
    # Penentuan FLRG
    for fuzzyfikasi in kumpulan_fuzzyfikasi:
        FLRG = []
        FLRG.append(fuzzyfikasi)
        relasiFLRG = []
        for i in FLRData_kumpulan:
            if fuzzyfikasi == i[0]: 
                relasiFLRG.append(i[1])
        print(list(dict.fromkeys(relasiFLRG)))
        relasiFLRG = list(dict.fromkeys(relasiFLRG))
        FLRG.append(relasiFLRG)
        perpus.append(FLRG)
    print(perpus)

def penentuanMatrix():
    # Penentuan untuk Matrix
    for fuzzyfikasi in kumpulan_fuzzyfikasi:
        FLRG = []
        FLRG.append(fuzzyfikasi)
        relasiFLRG = []
        for i in FLRData_kumpulan:
            if fuzzyfikasi == i[0]: 
                relasiFLRG.append(i[1])
        # print(list(dict.fromkeys(relasiFLRG)))
        # relasiFLRG = list(dict.fromkeys(relasiFLRG))
        FLRG.append(relasiFLRG)
        perpus_matrix.append(FLRG)
    print(perpus_matrix)

def hitungMatrix():
    # Perhitungan matriks
    for fuzzyfikasi in kumpulan_fuzzyfikasi:
        fuzzyfikasi_dan_relasi_flrgmc = []
        for i in perpus_matrix:
            kumpulan_relasi_flrgmc = []
            if i[0] == fuzzyfikasi:
                relasi_flrgmc = []
                for j in kumpulan_fuzzyfikasi:
                    if len(i[1]) == 0:
                        flrgmc = 0
                    else: 
                        flrgmc = i[1].count(j)/len(i[1])
                        satuan_flrgmc = []
                        satuan_flrgmc.append(j)
                        satuan_flrgmc.append(flrgmc)
                        relasi_flrgmc.append(satuan_flrgmc)
                        # print(flrgmc)
                        print(i[1].count(j))
                        # print(len(i[1]))
                # kumpulan_relasi_flrgmc.append(relasi_flrgmc)
                fuzzyfikasi_dan_relasi_flrgmc.append(i[0])
                fuzzyfikasi_dan_relasi_flrgmc.append(relasi_flrgmc)  
        array_of_matrix.append(fuzzyfikasi_dan_relasi_flrgmc)
    print(array_of_matrix)

def nilaiMatrix():
    # Perhitungan Nilai Matriks 
    for i in array_of_matrix:
        total_nilai_matrix = 0
        for j in i[1]:
            for k in kumpulan_nilai_tengah:
                if j[0] == k[0]:
                    nilai_matrix_relasi = j[1]*k[1]
                    total_nilai_matrix += nilai_matrix_relasi
        kelas_dan_nilai_matrix = []
        kelas_dan_nilai_matrix.append(i[0])
        kelas_dan_nilai_matrix.append(total_nilai_matrix)
        array_nilai_matrix.append(kelas_dan_nilai_matrix)
    print(array_nilai_matrix)
 
def nilaiFRLG():
    # Menghitung Nilai FLRG
    array_nilai_FLRG = []
    for buku in perpus:
        nilai_FLRG = 0
        banyak_FLRG = 0
        for flrg in buku[1]:
            for nilai_tengah in kumpulan_nilai_tengah:
                if(nilai_tengah[0] == flrg):
                    nilai_FLRG += nilai_tengah[1]
            banyak_FLRG += 1
        # print(nilai_FLRG)
        if banyak_FLRG == 0:
            nilai_FLRG = 0
        else:
            nilai_FLRG /= banyak_FLRG
        kelas_dan_nilai_FLRG = []
        kelas_dan_nilai_FLRG.append(buku[0])
        kelas_dan_nilai_FLRG.append(nilai_FLRG)
        array_nilai_FLRG.append(kelas_dan_nilai_FLRG)
    print(array_nilai_FLRG)
    # print("Nilai FLRG dari ", buku[0], " = ", str(nilai_FLRG))

def nilaiFLRGDataHistoris():
    #Menentukan Nilai FLRG untuk Data Historis
    for k in fuzzyfikasi_data_historis:
        for l in array_nilai_matrix:
            # hitung flrg
            if k[1] == l[0]:
                kurs_nilai_FLRG = []
                kurs_nilai_FLRG.append(k[0])
                kurs_nilai_FLRG.append(k[1])
                kurs_nilai_FLRG.append(l[1])
                # masukin ke peramalan
                if(len(array_kurs_nilai_FLRG) > 0):
                    kurs_nilai_FLRG_peramalan = []
                    nilai_peramalan = array_kurs_nilai_FLRG[len(array_kurs_nilai_FLRG)-1][2]
                    kurs_nilai_FLRG_peramalan.append(k[0])
                    kurs_nilai_FLRG_peramalan.append(k[1])
                    kurs_nilai_FLRG_peramalan.append(l[1])
                    kurs_nilai_FLRG_peramalan.append(nilai_peramalan)   
                    array_hasil_peramalan.append(kurs_nilai_FLRG_peramalan)
                array_kurs_nilai_FLRG.append(kurs_nilai_FLRG)

def hitungMAPE():
    #Menentukan MAPE
    array_hasil_mape = []
    for m in array_hasil_peramalan:
        hasil_mape = []
        nilai_mape = (abs(m[3] - m[0])/m[0]) * 100
        hasil_mape.append(m[0])
        hasil_mape.append(m[1])
        hasil_mape.append(m[2])
        hasil_mape.append(m[3])
        hasil_mape.append(nilai_mape)
        array_hasil_mape.append(hasil_mape)
        print("Kurs = ", m[0], " Fuzzyfikasi = ", m[1], " Nilai Matrix = ", m[2], " Hasil Peramalan = ", m[3], " Nilai MAPE = ", 
            nilai_mape)

    # Write to new excel
    # hasil_eksperimen = pd.DataFrame(array_hasil_peramalan, columns=['Kurs', 'Tanggal', 'Kurs', 'Fuzzyfikasi'])
    # with pd.ExcelWriter('D:\Kuliah\Semester8\TA\TA_KoTA305\TA\DataHistoris.xlsx', mode='a') as writer:
    #     FLR.to_excel(writer, sheet_name='Matrix Kurs Jual Sebelum (MC)', index=False)
    # print(data.values[0][0].astype(int))
    array_hasil_eksperimen = []
    j = 0
    for i in datav2.values:
        nomor = i[0]
        tanggal = str(i[1])
        kurs = 0
        fuzzyfikasi = ""
        nilai_FLRG = 0.0
        nilai_peramalan = 0.0
        nilai_mape = 0.0
        row_hasil_eksperimen = []
        # khusus row pertama
        if(j < 1):
            kurs = int(i[2])
            fuzzyfikasi = array_hasil_mape[j][1]
            nilai_FLRG = array_hasil_mape[j][2]
            row_hasil_eksperimen.append(nomor)
            row_hasil_eksperimen.append(tanggal)
            row_hasil_eksperimen.append(kurs)
            row_hasil_eksperimen.append(fuzzyfikasi)
            row_hasil_eksperimen.append(nilai_FLRG)
            
            array_hasil_eksperimen.append(row_hasil_eksperimen)
            
        elif j<=len(array_hasil_mape):
            kurs = array_hasil_mape[j-1][0]
            fuzzyfikasi = array_hasil_mape[j-1][1]
            nilai_FLRG = array_hasil_mape[j-1][2]
            nilai_peramalan = array_hasil_mape[j-1][3]
            nilai_mape = array_hasil_mape[j-1][4]
            row_hasil_eksperimen.append(nomor)
            row_hasil_eksperimen.append(tanggal)
            row_hasil_eksperimen.append(kurs)
            row_hasil_eksperimen.append(fuzzyfikasi)
            row_hasil_eksperimen.append(nilai_FLRG)
            row_hasil_eksperimen.append(nilai_peramalan)
            row_hasil_eksperimen.append(nilai_mape)
            
            array_hasil_eksperimen.append(row_hasil_eksperimen)
        
        j += 1
    
    # print(array_hasil_eksperimen)

    #Write To Excel
    
    excel_array_hasil_eksperimen = pd.DataFrame(array_hasil_eksperimen, columns=["No", "Tanggal", "Kurs", "Fuzzyfikasi", "Nilai Matrix", "Hasil Peramalan", "Nilai MAPE"])
    # excel_array_hasil_eksperimen.to_excel(sheet_name= 'Lembar1')
    with pd.ExcelWriter('D:\Kuliah\Semester8\TA\TA_KoTA305\TA\DataHistorisKursJualKetika.xlsx', mode='a') as writer:
        excel_array_hasil_eksperimen.to_excel(writer, sheet_name='Hasil Kurs 0000', index=False)
    
    
def main():
    intervalKelas()
    pengelompokkanFuzzyfikasiKurs()
    hitungFLR()
    hitungFLRG()
    penentuanMatrix()
    hitungMatrix()
    nilaiMatrix()
    nilaiFRLG()
    nilaiFLRGDataHistoris()
    hitungMAPE()
    
main()
    
# stop = timeit.default_timer()
# lama_eksekusi = stop - start
# print('Lama Waktu Eksekusi adalah: ', lama_eksekusi) 