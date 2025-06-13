import time
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==================== MAPPING ====================
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

# ==================== NORMALISASI ====================
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
        print(f"[DEBUG] Nilai mentah tanggal: {val}")
        if isinstance(val, pd.Timestamp):
            hasil = val.strftime("%Y-%m-%d")
            print(f"[DEBUG] Dari Timestamp: {hasil}")
            return hasil
        elif isinstance(val, (float, int)):
            date = pd.to_datetime("1899-12-30") + pd.to_timedelta(int(val), unit="D")
            hasil = date.strftime("%Y-%m-%d")
            print(f"[DEBUG] Dari serial Excel: {hasil}")
            return hasil
        elif isinstance(val, str):
            for fmt in ("%Y/%m/%d", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
                try:
                    dt = datetime.strptime(val.strip(), fmt)
                    hasil = dt.strftime("%Y-%m-%d")
                    print(f"[DEBUG] Dari string {fmt}: {hasil}")
                    return hasil
                except:
                    continue
            dt = pd.to_datetime(val, errors="coerce", dayfirst=True)
            if pd.notnull(dt):
                hasil = dt.strftime("%Y-%m-%d")
                print(f"[DEBUG] Dari fallback pd.to_datetime: {hasil}")
                return hasil
    except Exception as e:
        print(f"[!] Gagal parse tanggal '{val}': {e}")
    return ""

# ==================== SETUP ====================
options = Options()
options.headless = True
options.add_argument('--window-size=1920,1080')
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 10)
df = pd.read_excel("anggota.xlsx")
open("log_submit.txt", "w").close()

def log_to_file(text):
    with open("log_submit.txt", "a", encoding="utf-8") as f:
        f.write(text + "\n")

# ==================== FORM HANDLING ====================
def safe_send(name, value):
    try:
        input_el = wait.until(EC.presence_of_element_located((By.NAME, name)))
        input_el.clear()
        if pd.notna(value):
            if name == "tgl_register":
                driver.execute_script("arguments[0].value = arguments[1]", input_el, str(value))
                driver.execute_script("arguments[0].dispatchEvent(new Event('change'))", input_el)
                print(f"[DEBUG] Set 'tgl_register' dengan JS: {value}")
            else:
                input_el.send_keys(str(value))
    except Exception as e:
        log_to_file(f"[!] Gagal isi field: {name}")
        print(f"[!] Gagal isi field: {name} ({e})")

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
        print(f"[!] Gagal pilih option di: {name} = {value}")

# ==================== LOOP ====================
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
                msg = f"[!] Kelurahan dengan value '{kode_kel}' tidak tersedia di dropdown."
                log_to_file(msg)
                print(msg)
        else:
            msg = f"[!] Kelurahan '{row['kelurahan']}' tidak ditemukan di kelurahan_map"
            log_to_file(msg)
            print(msg)

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

# ==================== SELESAI ====================
driver.quit()
log_to_file("\n🎉 Semua data sudah diproses. Selesai!\n")
print("\n🎉 Semua data sudah diproses. Selesai!\n")
