import time
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import tempfile


jenis_kelamin_map = {"Laki-laki": "L", "Perempuan": "P"}
mondok_map = {"Belum Pernah": "N", "Pernah": "A"}
bersedia_text = ["Bersedia", "Tidak Bersedia"]

pekerjaan_valid = [
    "tidak_bekerja", "rumah_tangga", "pelajar_mahasiswa", "pensiunan",
    "pns", "tni_polri", "anggota_dpr", "kepala_daerah", "perangkat_desa",
    "karyawan", "dokter_perawat", "guru_dosen", "profesional",
    "arsitek_akuntan", "buruh", "tukang", "prt", "sopir_mekanik",
    "pedagang", "wirausaha", "petani_peternak", "usaha_lain",
    "seniman_wartawan", "ustadz", "paranormal", "lainnya"
]

dinas_map = {
    "Nahdlatul Ulama": 1, "Muslimat NU": 2, "GP Ansor": 3, "Fatayat NU": 4,
    "IPNU": 6, "IPPNU": 7, "JATMAN": 8, "PERGUNU": 9, "LDNU": 10, "Pagar Nusa": 11,
    "Ikatan Sarjana Nahdlatul Ulama": 12, "ISHARI NU": 13, "Lazis NU": 29,
    "LAKPESDAM": 46, "Lembaga Bahtsul Masail": 47, "LESBUMI": 48, "LPPNU": 45,
    "LKKNU": 44, "LPBINU": 49, "LTMNU": 50, "LP Ma'arif": 53, "JQH-NU": 52, "RMI-NU": 51
}

kelurahan_map = {
    "Lampar": 3309040001, "Dragan": 3309040002, "Karanganyar": 3309040003,
    "Jemowo": 3309040004, "Sumur": 3309040005, "Sangup": 3309040006,
    "Mriyan": 3309040007, "Lanjaran": 3309040008, "Karangkendal": 3309040009,
    "Keposong": 3309040010
}

def norm_kelamin(val):
    return jenis_kelamin_map.get(str(val).strip(), "")

def norm_mondok(val):
    return mondok_map.get(str(val).strip(), "")

def norm_pekerjaan(val):
    if pd.isna(val): return ""
    val = str(val).strip().lower().replace(" ", "_")
    return val if val in pekerjaan_valid else ""

def norm_dinas(val):
    return dinas_map.get(str(val).strip(), "")

def norm_kelurahan(val):
    return kelurahan_map.get(str(val).strip(), "")

def norm_bersedia(val):
    val = str(val).strip()
    return val if val in bersedia_text else "Tidak Bersedia"

def parse_tanggal(val):
    try:
        if isinstance(val, pd.Timestamp):
            return val.strftime("%Y-%m-%d")
        elif isinstance(val, (float, int)):
            date = pd.to_datetime("1899-12-30") + pd.to_timedelta(int(val), unit="D")
            return date.strftime("%Y-%m-%d")
        elif isinstance(val, str):
            for fmt in ("%Y/%m/%d", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
                try:
                    dt = datetime.strptime(val.strip(), fmt)
                    return dt.strftime("%Y-%m-%d")
                except:
                    continue
            dt = pd.to_datetime(val, errors="coerce", dayfirst=True)
            if pd.notnull(dt):
                return dt.strftime("%Y-%m-%d")
    except:
        return ""
    return ""

def run_bot(file_path):
    base = os.path.splitext(os.path.basename(file_path))[0]
    log_path = os.path.join("uploads", f"{base}_log.txt")
    open(log_path, "w").close()

    def log_to_file(text):
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(text + "\n")

    def safe_send(name, value):
        try:
            input_el = wait.until(EC.presence_of_element_located((By.NAME, name)))
            input_el.clear()
            if pd.notna(value):
                if name == "tgl_register":
                    driver.execute_script("arguments[0].value = arguments[1]", input_el, str(value))
                    driver.execute_script("arguments[0].dispatchEvent(new Event('change'))", input_el)
                else:
                    input_el.send_keys(str(value))
        except:
            log_to_file(f"[!] Gagal isi field: {name}")

    def safe_select(name, value, by_text=False):
        try:
            if pd.notna(value) and value != "":
                select_el = wait.until(EC.presence_of_element_located((By.NAME, name)))
                select_obj = Select(select_el)
                if by_text:
                    select_obj.select_by_visible_text(str(value))
                else:
                    select_obj.select_by_value(str(value))
        except:
            log_to_file(f"[!] Gagal pilih option di: {name} = {value}")

    options = Options()
    #options.headless = True
    #options.add_argument('--window-size=1920,1080')
    #options.add_argument('--window-size=1920,1080')
    #options.add_argument('--no-sandbox')
    #options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--headless=new')  # Lebih stabil di versi Chromium terbaru
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--window-size=1920,1080')

    # Gunakan user-data-dir unik agar tidak konflik
    temp_profile = tempfile.mkdtemp()
    options.add_argument(f'--user-data-dir={temp_profile}')
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 10)

    df = pd.read_excel(file_path)

    for index, row in df.iterrows():
        try:
            driver.get("https://enubo.nuboyolali.or.id/keanggotaan/daftar/N")

            safe_send("nama_muzaki", row["nama_muzaki"])
            safe_send("nik", row["nik"])
            safe_select("jenis_kelamin", norm_kelamin(row["jenis_kelamin"]))
            safe_send("tgl_register", parse_tanggal(row["tgl_register"]))
            safe_send("hp", row["hp"])
            safe_select("pekerjaan", norm_pekerjaan(row["pekerjaan"]))
            safe_select("mondok", norm_mondok(row["mondok"]))
            safe_send("email", row["email"])
            safe_send("npwp", row["npwp"])
            safe_send("alamat", row["alamat"])

            safe_select("kecamatan", "9471041")
            kode_kel = norm_kelurahan(row["kelurahan"])
            if kode_kel:
                try:
                    wait.until(lambda d: any(
                        option.get_attribute("value") == str(kode_kel)
                        for option in Select(d.find_element(By.NAME, "kelurahan")).options
                    ))
                    safe_select("kelurahan", str(kode_kel))
                except:
                    log_to_file(f"[!] Kelurahan dengan value '{kode_kel}' tidak tersedia.")
            else:
                log_to_file(f"[!] Kelurahan '{row['kelurahan']}' tidak ditemukan di map")

            safe_send("rt", row["rt"])
            safe_send("rw", row["rw"])
            safe_select("dinas", str(norm_dinas(row["dinas"])))
            safe_select("keanggotaan", row["keanggotaan"])
            safe_select("bersedia", norm_bersedia(row["bersedia"]), by_text=True)
            safe_select("keterangan", str(int(row["keterangan"])) if pd.notna(row["keterangan"]) else "")

            submit_btn = wait.until(EC.element_to_be_clickable((By.ID, "btnSubmit")))
            driver.execute_script("arguments[0].scrollIntoView(true);", submit_btn)
            time.sleep(0.5)
            submit_btn.click()
            time.sleep(2)

            msg = f"[✓] Baris ke-{index+1} ({row['nama_muzaki']}) berhasil disubmit."
            print(msg)
            log_to_file(msg)

        except Exception as e:
            msg = f"[✗] Baris ke-{index+1} gagal: {e}"
            print(msg)
            log_to_file(msg)

    driver.quit()
    log_to_file("\n🎉 Semua data sudah diproses. Selesai!\n")
    print("\n🎉 Semua data sudah diproses. Selesai!\n")

    return log_path, file_path
