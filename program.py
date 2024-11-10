# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog, messagebox
from docxtpl import DocxTemplate

def fill_template(values):
    template_path = filedialog.askopenfilename(title="Şablon Word dosyasını seçin", filetypes=[("Word dosyası", "*.docx")])
    if not template_path:
        return

    doc = DocxTemplate(template_path)

    doc.render(values)

    output_path = filedialog.asksaveasfilename(defaultextension=".docx", title="Kaydet", filetypes=[("Word dosyası", "*.docx")])
    if output_path:
        doc.save(output_path)
        messagebox.showinfo("Başarı", "Dosya başarıyla kaydedildi!")

def create_gui():
    root = tk.Tk()
    root.title("Yakalama Tutanağı Şablon Doldurma")

    left_frame = tk.Frame(root)
    left_frame.grid(row=0, column=0, padx=20, pady=10)
    
    right_frame = tk.Frame(root)
    right_frame.grid(row=0, column=1, padx=20, pady=10)

    entries = {}
    fields = {
        "tc": "T.C Kimlik Numarası",
        "ad_soyad": "Adı ve Soyadı",
        "baba_adi": "Baba Adı",
        "anne_adi": "Anne Adı",
        "dogum_yeri": "Doğum Yeri",
        "dogum_tarihi": "Doğum Tarihi",
        "nufus_kayitli_oldugu_yer": "Nüfusa Kayıtlı Olduğu Yer",
        "adres": "İkamet Adresi",
        "tel_no": "Telefon Numarası",
        "tespit_eden_kolluk": "Tespit Eden Kolluk Mensubu",
        "yakalama_tarihi": "Yakalama Tarihi",
        "yakalama_karar_saat": "Yakalama Karar Saati",
        "havayolu_firma": "Havayolu Firması",
        "sefer_num": "Sefer Numarası",
        "gelis_havaalanı": "Geliş Havaalanı",
        "kapi_no": "Kapı Numarası",
        "mahkeme_kararNo_kararTarih": "Mahkeme Karar No ve Karar Tarihi",
        "suc": "Suç",
        "yakalanma_saat": "Yakalanma Saati",
        "sahis_esyalari": "Üst Aramasında Çıkan Eşyalar",
        "tutanak_tarih": "Tutanak Tarihi",
        "tutanak_saat": "Tutanak Saati",
        "duzenleyen_memur": "Tutanak Düzenleyen Memur",
        "yazan_memur": "Tutanağı Yazan Memur",
        "arama_yapan": "Üst Araması Yapan Memur"
    }

    left_fields = list(fields.items())[:len(fields)//2]
    right_fields = list(fields.items())[len(fields)//2:]

    row = 0
    for key, label in left_fields:
        tk.Label(left_frame, text=label + ":").grid(row=row, column=0, sticky="e", padx=5, pady=5)
        entry = tk.Entry(left_frame, width=40)
        entry.grid(row=row, column=1, padx=5, pady=5)
        entries[key] = entry
        row += 1

    row = 0
    for key, label in right_fields:
        tk.Label(right_frame, text=label + ":").grid(row=row, column=0, sticky="e", padx=5, pady=5)
        entry = tk.Entry(right_frame, width=40)
        entry.grid(row=row, column=1, padx=5, pady=5)
        entries[key] = entry
        row += 1

    fill_button = tk.Button(root, text="Şablonu Doldur", command=lambda: fill_template({key: entry.get() for key, entry in entries.items()}))
    fill_button.grid(row=1, column=0, columnspan=2, pady=20)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
