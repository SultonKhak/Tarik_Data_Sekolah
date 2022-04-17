# import sys
from operator import le

import requests
from requests.adapters import HTTPAdapter
from requests.exceptions import ConnectionError
# import pandas as pd
import openpyxl
# import sharepoint_upload as su
# from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl import load_workbook


headers = {
    'user-agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36'
}

github_adapter = HTTPAdapter(max_retries=3)

session = requests.Session()
session.mount('https://dapo.kemdikbud.go.id/api', github_adapter)
workbook = Workbook()
sheet = workbook.active

def start_scrapping():

    number = 2

    url_provinsi = "https://dapo.kemdikbud.go.id/rekap/progres-sd?id_level_wilayah=0&kode_wilayah=000000&semester_id=20211"

    try:
        respone = session.get(url_provinsi, headers=headers)
    except ConnectionError as ce:
        print(ce)

    while respone.status_code < 200:
        respone = requests.get(url_provinsi)
        print(respone.status_code)

    if (respone.text == ''):
        data_provinsi = []
    else:
        data_provinsi = respone.json()

    if data_provinsi:
        for i in range(1,2):
            print(data_provinsi[i])
            url_kabupaten = "https://dapo.kemdikbud.go.id/rekap/progres-sd?id_level_wilayah=1&kode_wilayah="+ data_provinsi[i]["kode_wilayah"]+"&semester_id=20201"

            try:
                respone = session.get(url_kabupaten, headers=headers)
            except ConnectionError as ce:
                print(ce)

            while respone.status_code < 200:
                respone = requests.get(url_kabupaten)
                print(respone.status_code)

            if (respone.text == ''):
                data_kabupaten = []
            else:
                data_kabupaten = respone.json()

            if data_kabupaten:
                for j in range(len(data_kabupaten)):
                    print(data_kabupaten[j])
                    url_kecamatan = "https://dapo.kemdikbud.go.id/rekap/progres-sd?id_level_wilayah=2&kode_wilayah="+data_kabupaten[j]["kode_wilayah"]+"&semester_id=20201"

                    try:
                        respone = session.get(url_kecamatan, headers=headers)
                    except ConnectionError as ce:
                        print(ce)

                    while respone.status_code < 200:
                        respone = requests.get(url_kecamatan)
                        print(respone.status_code)

                    if (respone.text == ''):
                        data_kecamatan = []
                    else:
                        data_kecamatan = respone.json()

                    if data_kecamatan:
                        print("=============")
                        for k in range(len(data_kecamatan)):
                            print(data_kecamatan[k])
                            url_sekolah = "https://dapo.kemdikbud.go.id/rekap/progresSP-sd?id_level_wilayah=3&kode_wilayah="+data_kecamatan[k]["kode_wilayah"]+"&semester_id=20202"

                            try:
                                respone = session.get(url_sekolah, headers=headers)
                            except ConnectionError as ce:
                                print(ce)

                            while respone.status_code < 200:
                                respone = requests.get(url_sekolah)
                                print(respone.status_code)

                            if (respone.text == ''):
                                data_sekolah = []
                            else:
                                data_sekolah = respone.json()

                            if data_sekolah:
                                print("\n=============")
                                print(data_sekolah)

                                for l in range(len(data_sekolah)):
                                    print(data_sekolah[l])
                                    url_detail_sekolah = "https://dapo.kemdikbud.go.id/api/getHasilPencarian?keyword=" + str(data_sekolah[l]["npsn"])

                                    try:
                                        respone = session.get(url_detail_sekolah, headers=headers)
                                    except ConnectionError as ce:
                                        print(ce)

                                    while respone.status_code < 200:
                                        respone = requests.get(url_detail_sekolah)
                                        print(respone.status_code)

                                    if (respone.text == ''):
                                        data_detail_sekolah = []
                                    else:
                                        data_detail_sekolah = respone.json()

                                    if data_detail_sekolah:

                                        url_jumlah_siswa20202 = "https://dapo.kemdikbud.go.id/rekap/sekolahDetail?semester_id=20202&sekolah_id=" + data_detail_sekolah[0]["sekolah_id_enkrip"]
                                        url_jumlah_siswa20201 = "https://dapo.kemdikbud.go.id/rekap/sekolahDetail?semester_id=20201&sekolah_id=" + data_detail_sekolah[0]["sekolah_id_enkrip"]

                                        try:
                                            respone20202 = session.get(url_jumlah_siswa20202, headers=headers)

                                        except ConnectionError as ce:
                                            print(ce)

                                        try:
                                            respone20201 = session.get(url_jumlah_siswa20201, headers=headers)

                                        except ConnectionError as ce:
                                            print(ce)


                                        while respone20202.status_code < 200:
                                            respone20202 = requests.get(url_jumlah_siswa20202)
                                            print(respone20202.status_code)

                                        if (respone20202.text == ''):
                                            data_jumlah_siswa20202 = []
                                        else:
                                            data_jumlah_siswa20202 = respone20202.json()

                                        while respone20201.status_code < 200:
                                            respone20201 = requests.get(url_jumlah_siswa20201)
                                            print(respone20201.status_code)

                                        if (respone20201.text == ''):
                                            data_jumlah_siswa20201 = []
                                        else:
                                            data_jumlah_siswa20201 = respone20201.json()


                                        if data_jumlah_siswa20202:

                                            nama = data_sekolah[l]["nama"]
                                            npsn = data_sekolah[l]["npsn"]
                                            bentuk_pendidikan = data_sekolah[l]["bentuk_pendidikan_id"]
                                            status_sekolah = data_sekolah[l]["status_sekolah"]
                                            sekolah_id = data_sekolah[l]["sekolah_id"]
                                            sekolah_id_enkrip = data_sekolah[l]["sekolah_id_enkrip"]
                                            sinkron_terakhir = data_sekolah[l]["sinkron_terakhir"]
                                            alamat_jalan = data_detail_sekolah[0]["alamat_jalan"]
                                            kecamatan = data_detail_sekolah[0]["kecamatan"]
                                            kabupaten = data_detail_sekolah[0]["kabupaten"]
                                            propinsi = data_detail_sekolah[0]["propinsi"]
                                            guru_matematika20202 = data_jumlah_siswa20202[0]["guru_matematika"]
                                            guru_bahasa_indonesia20202 = data_jumlah_siswa20202[0]["guru_bahasa_indonesia"]
                                            guru_bahasa_inggris20202 = data_jumlah_siswa20202[0]["guru_bahasa_inggris"]
                                            guru_sejarah_indonesia20202 = data_jumlah_siswa20202[0]["guru_sejarah_indonesia"]
                                            guru_pkn20202 = data_jumlah_siswa20202[0]["guru_pkn"]
                                            guru_penjaskes20202 = data_jumlah_siswa20202[0]["guru_penjaskes"]
                                            guru_agama_budi_pekerti20202 = data_jumlah_siswa20202[0]["guru_agama_budi_pekerti"]
                                            guru_seni_budaya20202 = data_jumlah_siswa20202[0]["guru_seni_budaya"]
                                            ptk_laki20202 = data_jumlah_siswa20202[0]["ptk_laki"]
                                            ptk_perempuan20202 = data_jumlah_siswa20202[0]["ptk_perempuan"]
                                            ptk20202 = data_jumlah_siswa20202[0]["ptk"]
                                            pegawai_laki20202 = data_jumlah_siswa20202[0]["pegawai_laki"]
                                            pegawai_perempuan20202 = data_jumlah_siswa20202[0]["pegawai_perempuan"]
                                            pegawai20202 = data_jumlah_siswa20202[0]["pegawai"]

                                            if str(data_jumlah_siswa20202[0]["pd_kelas_10_laki"]) == "None":
                                                rombel_20202 = "0"
                                                guru_kelas_20202 = "0"
                                                pd_20202 = data_jumlah_siswa20202[0]["pd"]
                                                pd_laki_20202 = data_jumlah_siswa20202[0]["pd_laki"]
                                                kelas1_20202 = "0"
                                                kelas2_20202 = "0"
                                                kelas3_20202 = "0"
                                                kelas4_20202 = "0"
                                                kelas5_20202 = "0"
                                                kelas6_20202 = "0"
                                            else:
                                                rombel_20202 = data_jumlah_siswa20202[0]["rombel"]
                                                guru_kelas_20202 = data_jumlah_siswa20202[0]["guru_kelas"]
                                                pd_20202 = data_jumlah_siswa20202[0]["pd"]
                                                pd_laki_20202 = data_jumlah_siswa20202[0]["pd_laki"]
                                                kelas1_20202 = data_jumlah_siswa20202[0]["pd_kelas_1_laki"]+data_jumlah_siswa20202[0]["pd_kelas_1_perempuan"]
                                                kelas2_20202 = data_jumlah_siswa20202[0]["pd_kelas_2_laki"]+data_jumlah_siswa20202[0]["pd_kelas_2_perempuan"]
                                                kelas3_20202 = data_jumlah_siswa20202[0]["pd_kelas_3_laki"]+data_jumlah_siswa20202[0]["pd_kelas_3_perempuan"]
                                                kelas4_20202 = data_jumlah_siswa20202[0]["pd_kelas_4_laki"]+data_jumlah_siswa20202[0]["pd_kelas_4_perempuan"]
                                                kelas5_20202 = data_jumlah_siswa20202[0]["pd_kelas_5_laki"]+data_jumlah_siswa20202[0]["pd_kelas_5_perempuan"]
                                                kelas6_20202 = data_jumlah_siswa20202[0]["pd_kelas_6_laki"]+data_jumlah_siswa20202[0]["pd_kelas_6_perempuan"]

                                            pd_perempuan_20202 = data_jumlah_siswa20202[0]["pd_perempuan"]
                                            jumlah_kirim20202 = data_jumlah_siswa20202[0]["jumlah_kirim"]
                                            before_ruang_kelas20202 = data_jumlah_siswa20202[0]["before_ruang_kelas"]
                                            after_ruang_kelas20202 = data_jumlah_siswa20202[0]["after_ruang_kelas"]
                                            before_ruang_perpus20202 = data_jumlah_siswa20202[0]["before_ruang_perpus"]
                                            after_ruang_perpus20202 = data_jumlah_siswa20202[0]["after_ruang_perpus"]
                                            before_ruang_lab20202 = data_jumlah_siswa20202[0]["before_ruang_lab"]
                                            after_ruang_lab20202 = data_jumlah_siswa20202[0]["after_ruang_lab"]
                                            before_ruang_praktik20202 = data_jumlah_siswa20202[0]["before_ruang_praktik"]
                                            after_ruang_praktik20202 = data_jumlah_siswa20202[0]["after_ruang_praktik"]
                                            before_ruang_guru20202 = data_jumlah_siswa20202[0]["before_ruang_guru"]
                                            after_ruang_guru20202 = data_jumlah_siswa20202[0]["after_ruang_guru"]
                                            before_ruang_ibadah20202 = data_jumlah_siswa20202[0]["before_ruang_ibadah"]
                                            after_ruang_ibadah20202 = data_jumlah_siswa20202[0]["after_ruang_ibadah"]
                                            before_ruang_uks20202 = data_jumlah_siswa20202[0]["before_ruang_uks"]
                                            after_ruang_uks20202 = data_jumlah_siswa20202[0]["after_ruang_uks"]
                                            before_ruang_sirkulasi20202 = data_jumlah_siswa20202[0]["before_ruang_sirkulasi"]
                                            after_ruang_sirkulasi20202 = data_jumlah_siswa20202[0]["after_ruang_sirkulasi"]
                                            before_tempat_bermain_olahraga20202 = data_jumlah_siswa20202[0]["before_tempat_bermain_olahraga"]
                                            after_tempat_bermain_olahraga20202 = data_jumlah_siswa20202[0]["after_tempat_bermain_olahraga"]
                                            before_bangunan20202 = data_jumlah_siswa20202[0]["before_bangunan"]
                                            after_bangunan20202 = data_jumlah_siswa20202[0]["after_bangunan"]
                                            sumber_air20202 = data_jumlah_siswa20202[0]["sumber_air"]
                                            sumber_air_minum20202 = data_jumlah_siswa20202[0]["sumber_air_minum"]
                                            kecukupan_air_bersih20202 = data_jumlah_siswa20202[0]["kecukupan_air_bersih"]
                                        else:
                                            nama = ""
                                            npsn = ""
                                            bentuk_pendidikan = ""
                                            status_sekolah = ""
                                            sekolah_id = ""
                                            sekolah_id_enkrip = ""
                                            sinkron_terakhir = ""
                                            alamat_jalan = ""
                                            kecamatan = ""
                                            kabupaten = ""
                                            propinsi = ""
                                            rombel_20202 = "0"
                                            guru_kelas_20202 = "0"
                                            guru_matematika20202 = ""
                                            guru_bahasa_indonesia20202 = ""
                                            guru_bahasa_inggris20202 = ""
                                            guru_sejarah_indonesia20202 = ""
                                            guru_pkn20202 = ""
                                            guru_penjaskes20202 = ""
                                            guru_agama_budi_pekerti20202 = ""
                                            guru_seni_budaya20202 = ""
                                            ptk_laki20202 = ""
                                            ptk_perempuan20202 = ""
                                            ptk20202 = ""
                                            pegawai_laki20202 = ""
                                            pegawai_perempuan20202 = ""
                                            pegawai20202 = ""
                                            jumlah_kirim20202 = ""
                                            pd_20202 = "0"""
                                            pd_laki_20202 = "0"
                                            pd_perempuan_20202 = "0"
                                            before_ruang_kelas20202 = ""
                                            after_ruang_kelas20202 = ""
                                            before_ruang_perpus20202 = ""
                                            after_ruang_perpus20202 = ""
                                            before_ruang_lab20202 = ""
                                            after_ruang_lab20202 = ""
                                            before_ruang_praktik20202 = ""
                                            after_ruang_praktik20202 = ""
                                            before_ruang_guru20202 = ""
                                            after_ruang_guru20202 = ""
                                            before_ruang_ibadah20202 = ""
                                            after_ruang_ibadah20202 = ""
                                            before_ruang_uks20202 = ""
                                            after_ruang_uks20202 = ""
                                            before_ruang_sirkulasi20202 = ""
                                            after_ruang_sirkulasi20202 = ""
                                            before_tempat_bermain_olahraga20202 = ""
                                            after_tempat_bermain_olahraga20202 = ""
                                            before_bangunan20202 = ""
                                            after_bangunan20202 = ""
                                            sumber_air20202 = ""
                                            sumber_air_minum20202 = ""
                                            kecukupan_air_bersih20202 = ""
                                            kelas1_20202 = "0"
                                            kelas2_20202 = "0"
                                            kelas3_20202 = "0"
                                            kelas4_20202 = "0"
                                            kelas5_20202 = "0"
                                            kelas6_20202 = "0"


                                        if data_jumlah_siswa20201:
                                            guru_matematika20201 = data_jumlah_siswa20201[0]["guru_matematika"]
                                            guru_bahasa_indonesia20201 = data_jumlah_siswa20201[0]["guru_bahasa_indonesia"]
                                            guru_bahasa_inggris20201 = data_jumlah_siswa20201[0]["guru_bahasa_inggris"]
                                            guru_sejarah_indonesia20201 = data_jumlah_siswa20201[0]["guru_sejarah_indonesia"]
                                            guru_pkn20201 = data_jumlah_siswa20201[0]["guru_pkn"]
                                            guru_penjaskes20201 = data_jumlah_siswa20201[0]["guru_penjaskes"]
                                            guru_agama_budi_pekerti20201 = data_jumlah_siswa20201[0]["guru_agama_budi_pekerti"]
                                            guru_seni_budaya20201 = data_jumlah_siswa20202[0]["guru_seni_budaya"]
                                            ptk_laki20201 = data_jumlah_siswa20201[0]["ptk_laki"]
                                            ptk_perempuan20201 = data_jumlah_siswa20201[0]["ptk_perempuan"]
                                            ptk20201 = data_jumlah_siswa20201[0]["ptk"]
                                            pegawai_laki20201 = data_jumlah_siswa20201[0]["pegawai_laki"]
                                            pegawai_perempuan20201 = data_jumlah_siswa20201[0]["pegawai_perempuan"]
                                            pegawai20201 = data_jumlah_siswa20201[0]["pegawai"]

                                            if str(data_jumlah_siswa20201[0]["pd_kelas_10_laki"]) == "None":
                                                rombel_20201 = "0"
                                                guru_kelas_20201 = "0"
                                                pd_20201 = data_jumlah_siswa20201[0]["pd"]
                                                pd_laki_20201 = data_jumlah_siswa20201[0]["pd_laki"]
                                                kelas1_20201 = "0"
                                                kelas2_20201 = "0"
                                                kelas3_20201 = "0"
                                                kelas4_20201 = "0"
                                                kelas5_20201 = "0"
                                                kelas6_20201 = "0"
                                            else:
                                                rombel_20201 = data_jumlah_siswa20201[0]["rombel"]
                                                guru_kelas_20201 = data_jumlah_siswa20201[0]["guru_kelas"]
                                                pd_20201 = data_jumlah_siswa20201[0]["pd"]
                                                pd_laki_20201 = data_jumlah_siswa20201[0]["pd_laki"]
                                                kelas1_20201 = data_jumlah_siswa20201[0]["pd_kelas_1_laki"]+data_jumlah_siswa20201[0]["pd_kelas_1_perempuan"]
                                                kelas2_20201 = data_jumlah_siswa20201[0]["pd_kelas_2_laki"]+data_jumlah_siswa20201[0]["pd_kelas_2_perempuan"]
                                                kelas3_20201 = data_jumlah_siswa20201[0]["pd_kelas_3_laki"]+data_jumlah_siswa20201[0]["pd_kelas_3_perempuan"]
                                                kelas4_20201 = data_jumlah_siswa20201[0]["pd_kelas_4_laki"]+data_jumlah_siswa20201[0]["pd_kelas_4_perempuan"]
                                                kelas5_20201 = data_jumlah_siswa20201[0]["pd_kelas_5_laki"]+data_jumlah_siswa20201[0]["pd_kelas_5_perempuan"]
                                                kelas6_20201 = data_jumlah_siswa20201[0]["pd_kelas_6_laki"]+data_jumlah_siswa20201[0]["pd_kelas_6_perempuan"]

                                            pd_perempuan_20201 = data_jumlah_siswa20201[0]["pd_perempuan"]
                                            jumlah_kirim20201 = data_jumlah_siswa20201[0]["jumlah_kirim"]
                                            before_ruang_kelas20201 = data_jumlah_siswa20201[0]["before_ruang_kelas"]
                                            after_ruang_kelas20201 = data_jumlah_siswa20201[0]["after_ruang_kelas"]
                                            before_ruang_perpus20201 = data_jumlah_siswa20201[0]["before_ruang_perpus"]
                                            after_ruang_perpus20201 = data_jumlah_siswa20201[0]["after_ruang_perpus"]
                                            before_ruang_lab20201 = data_jumlah_siswa20201[0]["before_ruang_lab"]
                                            after_ruang_lab20201 = data_jumlah_siswa20202[0]["after_ruang_lab"]
                                            before_ruang_praktik20201 = data_jumlah_siswa20201[0]["before_ruang_praktik"]
                                            after_ruang_praktik20201 = data_jumlah_siswa20201[0]["after_ruang_praktik"]
                                            before_ruang_guru20201 = data_jumlah_siswa20201[0]["before_ruang_guru"]
                                            after_ruang_guru20201 = data_jumlah_siswa20201[0]["after_ruang_guru"]
                                            before_ruang_ibadah20201 = data_jumlah_siswa20201[0]["before_ruang_ibadah"]
                                            after_ruang_ibadah20201 = data_jumlah_siswa20201[0]["after_ruang_ibadah"]
                                            before_ruang_uks20201 = data_jumlah_siswa20201[0]["before_ruang_uks"]
                                            after_ruang_uks20201 = data_jumlah_siswa20201[0]["after_ruang_uks"]
                                            before_ruang_sirkulasi20201 = data_jumlah_siswa20201[0]["before_ruang_sirkulasi"]
                                            after_ruang_sirkulasi20201 = data_jumlah_siswa20201[0]["after_ruang_sirkulasi"]
                                            before_tempat_bermain_olahraga20201 = data_jumlah_siswa20201[0]["before_tempat_bermain_olahraga"]
                                            after_tempat_bermain_olahraga20201 = data_jumlah_siswa20201[0]["after_tempat_bermain_olahraga"]
                                            before_bangunan20201 = data_jumlah_siswa20201[0]["before_bangunan"]
                                            after_bangunan20201 = data_jumlah_siswa20201[0]["after_bangunan"]
                                            sumber_air20201 = data_jumlah_siswa20201[0]["sumber_air"]
                                            sumber_air_minum20201 = data_jumlah_siswa20201[0]["sumber_air_minum"]
                                            kecukupan_air_bersih20201 = data_jumlah_siswa20201[0]["kecukupan_air_bersih"]
                                        else:
                                            rombel_20201 = "0"
                                            guru_kelas_20201 = "0"
                                            guru_matematika20201 = ""
                                            guru_bahasa_indonesia20201 = ""
                                            guru_bahasa_inggris20201 = ""
                                            guru_sejarah_indonesia20201 = ""
                                            guru_pkn20201 = ""
                                            guru_penjaskes20201 = ""
                                            guru_agama_budi_pekerti20201 = ""
                                            guru_seni_budaya20201 = ""
                                            ptk_laki20201 = ""
                                            ptk_perempuan20201 = ""
                                            ptk20201 = ""
                                            pegawai_laki20201 = ""
                                            pegawai_perempuan20201 = ""
                                            pegawai20201 = ""
                                            jumlah_kirim20201 = ""
                                            pd_20201 = "0"""
                                            pd_laki_20201 = "0"
                                            pd_perempuan_20201 = "0"
                                            before_ruang_kelas20201 = ""
                                            after_ruang_kelas20201 = ""
                                            before_ruang_perpus20201 = ""
                                            after_ruang_perpus20201 = ""
                                            before_ruang_lab20201 = ""
                                            after_ruang_lab20201 = ""
                                            before_ruang_praktik20201 = ""
                                            after_ruang_praktik20201 = ""
                                            before_ruang_guru20201 = ""
                                            after_ruang_guru20201 = ""
                                            before_ruang_ibadah20201 = ""
                                            after_ruang_ibadah20201 = ""
                                            before_ruang_uks20201 = ""
                                            after_ruang_uks20201 = ""
                                            before_ruang_sirkulasi20201 = ""
                                            after_ruang_sirkulasi20201 = ""
                                            before_tempat_bermain_olahraga20201 = ""
                                            after_tempat_bermain_olahraga20201 = ""
                                            before_bangunan20201 = ""
                                            after_bangunan20201 = ""
                                            sumber_air20201 = ""
                                            sumber_air_minum20201 = ""
                                            kecukupan_air_bersih20201 = ""
                                            kelas1_20201 = "0"
                                            kelas2_20201 = "0"
                                            kelas3_20201 = "0"
                                            kelas4_20201 = "0"
                                            kelas5_20201 = "0"
                                            kelas6_20201 = "0"


                                        sheet['A' + '1'].value = "nama"
                                        sheet['B' + '1'].value = "npsn"
                                        sheet['C' + '1'].value = "bentuk_pendidikan"
                                        sheet['D' + '1'].value = "status_sekolah"
                                        sheet['E' + '1'].value = "sekolah_id"
                                        sheet['F' + '1'].value = "sekolah_id_enkrip"
                                        sheet['G' + '1'].value = "sinkron_terakhir"
                                        sheet['H' + '1'].value = "alamat_jalan"
                                        sheet['I' + '1'].value = "kecamatan"
                                        sheet['J' + '1'].value = "kabupaten"
                                        sheet['K' + '1'].value = "propinsi"
                                        sheet['M' + '1'].value = "guru_kelas_20201"
                                        sheet['N' + '1'].value = "guru_kelas_20202"
                                        sheet['O' + '1'].value = "guru_matematika20201"
                                        sheet['P' + '1'].value = "guru_matematika20202"
                                        sheet['Q' + '1'].value = "guru_bahasa_indonesia20201"
                                        sheet['R' + '1'].value = "guru_bahasa_indonesia20202"
                                        sheet['S' + '1'].value = "guru_bahasa_inggris20201"
                                        sheet['T' + '1'].value = "guru_bahasa_inggris20202"
                                        sheet['U' + '1'].value = "guru_sejarah_indonesia20201"
                                        sheet['V' + '1'].value = "guru_sejarah_indonesia20202"
                                        sheet['W' + '1'].value = "guru_pkn20201"
                                        sheet['X' + '1'].value = "guru_pkn20202"
                                        sheet['Y' + '1'].value = "guru_penjaskes20201"
                                        sheet['Z' + '1'].value = "guru_penjaskes20202"
                                        sheet['AA' + '1'].value = "guru_agama_budi_pekerti20201"
                                        sheet['AB' + '1'].value = "guru_agama_budi_pekerti20202"
                                        sheet['AC' + '1'].value = "guru_seni_budaya20201"
                                        sheet['AD' + '1'].value = "guru_seni_budaya20202"
                                        sheet['AE' + '1'].value = "ptk_laki20201"
                                        sheet['AF' + '1'].value = "ptk_laki20202"
                                        sheet['AG' + '1'].value = "ptk_perempuan20201"
                                        sheet['AH' + '1'].value = "ptk_perempuan20202"
                                        sheet['AI' + '1'].value = "ptk20201"
                                        sheet['AJ' + '1'].value = "ptk20202"
                                        sheet['AK' + '1'].value = "pegawai_laki20201"
                                        sheet['AL' + '1'].value = "pegawai_laki20202"
                                        sheet['AM' + '1'].value = "pegawai_perempuan20201"
                                        sheet['AN' + '1'].value = "pegawai_perempuan20202"
                                        sheet['AO' + '1'].value = "pegawai20201"
                                        sheet['AP' + '1'].value = "pegawai20202"
                                        sheet['AQ' + '1'].value = "before_ruang_kelas20201"
                                        sheet['AR' + '1'].value = "before_ruang_kelas20202"
                                        sheet['AS' + '1'].value = "after_ruang_kelas20201"
                                        sheet['AT' + '1'].value = "after_ruang_kelas20202"
                                        sheet['AU' + '1'].value = "before_ruang_perpus20201"
                                        sheet['AV' + '1'].value = "before_ruang_perpus20202"
                                        sheet['AW' + '1'].value = "after_ruang_perpus20201"
                                        sheet['AX' + '1'].value = "after_ruang_perpus20202"
                                        sheet['AY' + '1'].value = "before_ruang_lab20201"
                                        sheet['AZ' + '1'].value = "before_ruang_lab20202"
                                        sheet['BA' + '1'].value = "after_ruang_lab20201"
                                        sheet['BB' + '1'].value = "after_ruang_lab20202"
                                        sheet['BC' + '1'].value = "before_ruang_praktik20201"
                                        sheet['BD' + '1'].value = "before_ruang_praktik20202"
                                        sheet['BE' + '1'].value = "after_ruang_praktik20201"
                                        sheet['BF' + '1'].value = "after_ruang_praktik20202"
                                        sheet['BG' + '1'].value = "before_ruang_guru20201"
                                        sheet['BH' + '1'].value = "before_ruang_guru20202"
                                        sheet['BI' + '1'].value = "after_ruang_guru20201"
                                        sheet['BJ' + '1'].value = "after_ruang_guru20202"
                                        sheet['BK' + '1'].value = "before_ruang_ibadah20201"
                                        sheet['BL' + '1'].value = "before_ruang_ibadah20202"
                                        sheet['BM' + '1'].value = "after_ruang_ibadah20201"
                                        sheet['BN' + '1'].value = "after_ruang_ibadah20202"
                                        sheet['BO' + '1'].value = "before_ruang_uks20201"
                                        sheet['BP' + '1'].value = "before_ruang_uks20202"
                                        sheet['BQ' + '1'].value = "after_ruang_uks20201"
                                        sheet['BR' + '1'].value = "after_ruang_uks20202"
                                        sheet['BS' + '1'].value = "before_ruang_sirkulasi20201"
                                        sheet['BT' + '1'].value = "before_ruang_sirkulasi20202"
                                        sheet['BU' + '1'].value = "after_ruang_sirkulasi20201"
                                        sheet['BV' + '1'].value = "after_ruang_sirkulasi20202"
                                        sheet['BW' + '1'].value = "before_tempat_bermain_olahraga20201"
                                        sheet['BX' + '1'].value = "before_tempat_bermain_olahraga20202"
                                        sheet['BY' + '1'].value = "after_tempat_bermain_olahraga20201"
                                        sheet['BZ' + '1'].value = "after_tempat_bermain_olahraga20202"
                                        sheet['CA' + '1'].value = "before_bangunan20201"
                                        sheet['CB' + '1'].value = "before_bangunan20202"
                                        sheet['CC' + '1'].value = "after_bangunan20201"
                                        sheet['CD' + '1'].value = "after_bangunan20202"
                                        sheet['CE' + '1'].value = "sumber_air20201"
                                        sheet['CF' + '1'].value = "sumber_air20202"
                                        sheet['CG' + '1'].value = "sumber_air_minum20201"
                                        sheet['CH' + '1'].value = "sumber_air_minum20202"
                                        sheet['CI' + '1'].value = "kecukupan_air_bersih20201"
                                        sheet['CJ' + '1'].value = "kecukupan_air_bersih20202"
                                        sheet['CK' + '1'].value = "jumlah_kirim20201"
                                        sheet['CL' + '1'].value = "jumlah_kirim20202"
                                        sheet['CM' + '1'].value = "rombel_20201"
                                        sheet['CN' + '1'].value = "rombel_20202"
                                        sheet['CO' + '1'].value = "pd_20201"
                                        sheet['CP' + '1'].value = "pd_20202"
                                        sheet['CQ' + '1'].value = "pd_laki_20201"
                                        sheet['CR' + '1'].value = "pd_laki_20202"
                                        sheet['CS' + '1'].value = "pd_perempuan_20201"
                                        sheet['CT' + '1'].value = "pd_perempuan_20202"
                                        sheet['CU' + '1'].value = "kelas1_20201"
                                        sheet['CV' + '1'].value = "kelas1_20202"
                                        sheet['CW' + '1'].value = "kelas2_20201"
                                        sheet['CX' + '1'].value = "kelas2_20202"
                                        sheet['CY' + '1'].value = "kelas3_20201"
                                        sheet['CZ' + '1'].value = "kelas3_20202"
                                        sheet['DA' + '1'].value = "kelas4_20201"
                                        sheet['DB' + '1'].value = "kelas4_20202"
                                        sheet['DC' + '1'].value = "kelas5_20201"
                                        sheet['DD' + '1'].value = "kelas5_20202"
                                        sheet['DE' + '1'].value = "kelas6_20201"
                                        sheet['DF' + '1'].value = "kelas6_20202"





                                        sheet['A' + str(number)].value = nama
                                        sheet['B' + str(number)].value = npsn
                                        sheet['C' + str(number)].value = bentuk_pendidikan
                                        sheet['D' + str(number)].value = status_sekolah
                                        sheet['E' + str(number)].value = sekolah_id
                                        sheet['F' + str(number)].value = sekolah_id_enkrip
                                        sheet['G' + str(number)].value = sinkron_terakhir
                                        sheet['H' + str(number)].value = alamat_jalan
                                        sheet['I' + str(number)].value = kecamatan
                                        sheet['J' + str(number)].value = kabupaten
                                        sheet['K' + str(number)].value = propinsi
                                        sheet['M' + str(number)].value = guru_kelas_20201
                                        sheet['N' + str(number)].value = guru_kelas_20202
                                        sheet['O' + str(number)].value = guru_matematika20201
                                        sheet['P' + str(number)].value = guru_matematika20202
                                        sheet['Q' + str(number)].value = guru_bahasa_indonesia20201
                                        sheet['R' + str(number)].value = guru_bahasa_indonesia20202
                                        sheet['S' + str(number)].value = guru_bahasa_inggris20201
                                        sheet['T' + str(number)].value = guru_bahasa_inggris20202
                                        sheet['U' + str(number)].value = guru_sejarah_indonesia20201
                                        sheet['V' + str(number)].value = guru_sejarah_indonesia20202
                                        sheet['W' + str(number)].value = guru_pkn20201
                                        sheet['X' + str(number)].value = guru_pkn20202
                                        sheet['Y' + str(number)].value = guru_penjaskes20201
                                        sheet['Z' + str(number)].value = guru_penjaskes20202
                                        sheet['AA' + str(number)].value = guru_agama_budi_pekerti20201
                                        sheet['AB' + str(number)].value = guru_agama_budi_pekerti20202
                                        sheet['AC' + str(number)].value = guru_seni_budaya20201
                                        sheet['AD' + str(number)].value = guru_seni_budaya20202
                                        sheet['AE' + str(number)].value = ptk_laki20201
                                        sheet['AF' + str(number)].value = ptk_laki20202
                                        sheet['AG' + str(number)].value = ptk_perempuan20201
                                        sheet['AH' + str(number)].value = ptk_perempuan20202
                                        sheet['AI' + str(number)].value = ptk20201
                                        sheet['AJ' + str(number)].value = ptk20202
                                        sheet['AK' + str(number)].value = pegawai_laki20201
                                        sheet['AL' + str(number)].value = pegawai_laki20202
                                        sheet['AM' + str(number)].value = pegawai_perempuan20201
                                        sheet['AN' + str(number)].value = pegawai_perempuan20202
                                        sheet['AO' + str(number)].value = pegawai20201
                                        sheet['AP' + str(number)].value = pegawai20202
                                        sheet['AQ' + str(number)].value = before_ruang_kelas20201
                                        sheet['AR' + str(number)].value = before_ruang_kelas20202
                                        sheet['AS' + str(number)].value = after_ruang_kelas20201
                                        sheet['AT' + str(number)].value = after_ruang_kelas20202
                                        sheet['AU' + str(number)].value = before_ruang_perpus20201
                                        sheet['AV' + str(number)].value = before_ruang_perpus20202
                                        sheet['AW' + str(number)].value = after_ruang_perpus20201
                                        sheet['AX' + str(number)].value = after_ruang_perpus20202
                                        sheet['AY' + str(number)].value = before_ruang_lab20201
                                        sheet['AZ' + str(number)].value = before_ruang_lab20202
                                        sheet['BA' + str(number)].value = after_ruang_lab20201
                                        sheet['BB' + str(number)].value = after_ruang_lab20202
                                        sheet['BC' + str(number)].value = before_ruang_praktik20201
                                        sheet['BD' + str(number)].value = before_ruang_praktik20202
                                        sheet['BE' + str(number)].value = after_ruang_praktik20201
                                        sheet['BF' + str(number)].value = after_ruang_praktik20202
                                        sheet['BG' + str(number)].value = before_ruang_guru20201
                                        sheet['BH' + str(number)].value = before_ruang_guru20202
                                        sheet['BI' + str(number)].value = after_ruang_guru20201
                                        sheet['BJ' + str(number)].value = after_ruang_guru20202
                                        sheet['BK' + str(number)].value = before_ruang_ibadah20201
                                        sheet['BL' + str(number)].value = before_ruang_ibadah20202
                                        sheet['BM' + str(number)].value = after_ruang_ibadah20201
                                        sheet['BN' + str(number)].value = after_ruang_ibadah20202
                                        sheet['BO' + str(number)].value = before_ruang_uks20201
                                        sheet['BP' + str(number)].value = before_ruang_uks20202
                                        sheet['BQ' + str(number)].value = after_ruang_uks20201
                                        sheet['BR' + str(number)].value = after_ruang_uks20202
                                        sheet['BS' + str(number)].value = before_ruang_sirkulasi20201
                                        sheet['BT' + str(number)].value = before_ruang_sirkulasi20202
                                        sheet['BU' + str(number)].value = after_ruang_sirkulasi20201
                                        sheet['BV' + str(number)].value = after_ruang_sirkulasi20202
                                        sheet['BW' + str(number)].value = before_tempat_bermain_olahraga20201
                                        sheet['BX' + str(number)].value = before_tempat_bermain_olahraga20202
                                        sheet['BY' + str(number)].value = after_tempat_bermain_olahraga20201
                                        sheet['BZ' + str(number)].value = after_tempat_bermain_olahraga20202
                                        sheet['CA' + str(number)].value = before_bangunan20201
                                        sheet['CB' + str(number)].value = before_bangunan20202
                                        sheet['CC' + str(number)].value = after_bangunan20201
                                        sheet['CD' + str(number)].value = after_bangunan20202
                                        sheet['CE' + str(number)].value = sumber_air20201
                                        sheet['CF' + str(number)].value = sumber_air20202
                                        sheet['CG' + str(number)].value = sumber_air_minum20201
                                        sheet['CH' + str(number)].value = sumber_air_minum20202
                                        sheet['CI' + str(number)].value = kecukupan_air_bersih20201
                                        sheet['CJ' + str(number)].value = kecukupan_air_bersih20202
                                        sheet['CK' + str(number)].value = jumlah_kirim20201
                                        sheet['CL' + str(number)].value = jumlah_kirim20202
                                        sheet['CM' + str(number)].value = rombel_20201
                                        sheet['CN' + str(number)].value = rombel_20202
                                        sheet['CO' + str(number)].value = pd_20201
                                        sheet['CP' + str(number)].value = pd_20202
                                        sheet['CQ' + str(number)].value = pd_laki_20201
                                        sheet['CR' + str(number)].value = pd_laki_20202
                                        sheet['CS' + str(number)].value = pd_perempuan_20201
                                        sheet['CT' + str(number)].value = pd_perempuan_20202
                                        sheet['CU' + str(number)].value = kelas1_20201
                                        sheet['CV' + str(number)].value = kelas1_20202
                                        sheet['CW' + str(number)].value = kelas2_20201
                                        sheet['CX' + str(number)].value = kelas2_20202
                                        sheet['CY' + str(number)].value = kelas3_20201
                                        sheet['CZ' + str(number)].value = kelas3_20202
                                        sheet['DA' + str(number)].value = kelas4_20201
                                        sheet['DB' + str(number)].value = kelas4_20202
                                        sheet['DC' + str(number)].value = kelas5_20201
                                        sheet['DD' + str(number)].value = kelas5_20202
                                        sheet['DE' + str(number)].value = kelas6_20201
                                        sheet['DF' + str(number)].value = kelas6_20202


                                        number = number + 1

                            # excelKecamatan = str(data_kecamatan[k]["nama"])+".xlsx"
                            # workbook.save(filename= excelKecamatan)

                    excelKabupaten = str(data_kabupaten[j]["nama"])+".xlsx"
                    workbook.save(filename= excelKabupaten)

            excel = str(data_provinsi[i]["nama"])+".xlsx"
            workbook.save(filename= excel)

start_scrapping()

