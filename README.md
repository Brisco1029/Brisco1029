import tkinter as tk
from tkinter import *
from tkinter.ttk import Combobox, Separator
from tkinter import ttk, messagebox
import datetime, time
from tkcalendar import DateEntry
import pathlib
import openpyxl
from openpyxl import Workbook, load_workbook

# Create main window with Title and window size (unresizeable)
window = Tk()
window.title("Citizen Registration System 2023")
window.iconbitmap('images/Aplication_icon.ico')
window.geometry('1070x630+125+25')
window.resizable(0,0)


# Create Resource
font1=('Times','14')
file = Workbook()

# Create Userform Title 
lbl_userform = Label(window,text="  Form Input Data ",height=1,font=('Times','14','bold'),bg='cyan',relief=RIDGE,anchor=W,borderwidth=2)
lbl_userform.pack(side=TOP,fill=X)

file=pathlib.Path('Citizen_Registration_System.xlsx')
if file.exists():
    pass
else:
    file=Workbook()
    sheet=file.active
    sheet['A2']="Serial No."
    sheet['B2']="No. KK"
    sheet['C2']="No. NIK"
    sheet['D2']="Nama Lengkap"
    sheet['E2']="Jenis Kelamin"
    sheet['F2']="Alamat"
    sheet['G2']="Tempat Lahir"
    sheet['H2']="Tanggal Lahir"
    sheet['I2']="Agama"
    sheet['J2']="Golongan Darah"
    sheet['K2']="Pendidikan Akhir"
    sheet['L2']="Jenis Pekerjaan"
    sheet['M2']="Status Perkawinan"
    sheet['N2']="Tanggal Perkawinan"
    sheet['O2']="Hubungan Keluarga"
    sheet['P2']="Kewarganegaraan"
    sheet['Q2']="No. Pasport"
    sheet['R2']="No. KITAP"
    sheet['S2']="Nama Ayah"
    sheet['T2']="Nama Ibu"
    sheet['U2']="Status Domisili"
    sheet['V2']="Akte Kelahiran"
    sheet['W2']="No. Akte Kelahiran"
    sheet['X2']="Akte Kematian"
    sheet['Y2']="No. Akte Kematian"
    sheet['Z2']="Fotocopy Kartu Keluarga"
    file.save('Citizen_Registration_System.xlsx')

# Create define edit Data
def edit():
    pass

# Create define Clear Data    
def clear():
    kk_var.set('')
    nik_var.set('')
    nama_var.set('')
    gen_var.set(' Laki-laki ')
    alamat_var.set('')
    goldar_var.set(' Pilih Golongan Darah ')
    kota_var.set('')
    martial_var.set(' Pilih Status Perkawinan ')
    educ_var.set(' Pendidikan Akhir ')
    job_var.set(' Pilih Jenis Pekerjaan ')
    religi_var.set(' Pilih Agama / Keyakinan ')
    fam_var.set(' Pilih Hubungan Keluarga ')
    nat_var.set(' Pilih Kewarganegaraan ')
    pas_var.set('')
    kit_var.set('')
    ayah_var.set('')
    ibu_var.set('')
    residence_var.set(' Pilih Status Domisili  ')
    fc_kk_var.set(' Pilih Ada / Tidak  ')
    ivar.set(clear)
    ivari.set(clear)
    akte_lahir_var.set('')
    akte_mati_var.set('')

# Create define Submit/Save
def submit():
    ser_num=no_var.get()
    kk=kk_var.get()
    nik=nik_var.get()
    nama=nama_var.get()
    gen=gen_var.get()
    alamat=alamat_var.get()
    goldar=goldar_var.get()
    kota=kota_var.get()
    birth_date=birth_date_entry.get()
    martial=martial_var.get()
    martial_date=martial_date_entry.get()
    educ=educ_var.get()
    job=job_var.get()
    religi=religi_var.get()
    fam=fam_var.get()
    nat=nat_var.get()
    pas=pas_var.get()
    kit=kit_var.get()
    ayah=ayah_var.get()
    ibu=ibu_var.get()
    residence=residence_var.get()
    fc_kk=fc_kk_var.get()
    akte_lahir=akte_lahir_check.get()
    akte_mati=akte_mati_check.get()
    no_akte_lahir=akte_lahir_var.get()
    no_akte_mati=akte_mati_var.get()

    if (not kk or not nik or not nama or not gen or not alamat or not goldar or not kota or not birth_date or not martial or not martial_date or not educ or not job or not religi or not fam or not nat or not pas or not kit or not ayah or not ibu or not residence or not fc_kk or not no_akte_lahir or not no_akte_mati):
        messagebox.showerror("Error", "Please fill in all fields")
        return

    file=openpyxl.load_workbook('Citizen_Registration_System.xlsx')
    sheet=file.active
    sheet.cell(column=1,row=sheet.max_row+1,value=ser_num)
    sheet.cell(column=2,row=sheet.max_row,value=kk)
    sheet.cell(column=3,row=sheet.max_row,value=nik)        
    sheet.cell(column=4,row=sheet.max_row,value=nama)
    sheet.cell(column=5,row=sheet.max_row,value=gen)
    sheet.cell(column=6,row=sheet.max_row,value=alamat)
    sheet.cell(column=7,row=sheet.max_row,value=goldar)
    sheet.cell(column=8,row=sheet.max_row,value=kota)
    sheet.cell(column=9,row=sheet.max_row,value=birth_date)
    sheet.cell(column=10,row=sheet.max_row,value=martial)
    sheet.cell(column=11,row=sheet.max_row,value=martial_date)        
    sheet.cell(column=12,row=sheet.max_row,value=educ)
    sheet.cell(column=13,row=sheet.max_row,value=job)
    sheet.cell(column=14,row=sheet.max_row,value=religi)
    sheet.cell(column=15,row=sheet.max_row,value=fam)
    sheet.cell(column=16,row=sheet.max_row,value=nat)
    sheet.cell(column=17,row=sheet.max_row,value=pas)
    sheet.cell(column=18,row=sheet.max_row,value=kit)        
    sheet.cell(column=19,row=sheet.max_row,value=ayah)
    sheet.cell(column=20,row=sheet.max_row,value=ibu)
    sheet.cell(column=21,row=sheet.max_row,value=residence)
    sheet.cell(column=22,row=sheet.max_row,value=fc_kk)
    sheet.cell(column=23,row=sheet.max_row,value=akte_lahir)
    sheet.cell(column=24,row=sheet.max_row,value=akte_mati)
    sheet.cell(column=25,row=sheet.max_row,value=no_akte_lahir)
    sheet.cell(column=26,row=sheet.max_row,value=no_akte_mati)
    file.save(r'Citizen_Registration_System.xlsx')
    messagebox.showinfo(title='CRS Info',message='Data sudah disimpan dalam Database')

# Clear field after saving entry
    kk_var.set('')
    nik_var.set('')
    nama_var.set('')
    gen_var.set(' Laki-laki ')
    alamat_var.set('')
    goldar_var.set(' Pilih Golongan Darah ')
    kota_var.set('')
    martial_var.set(' Pilih Status Perkawinan ')
    educ_var.set(' Pendidikan Akhir ')
    job_var.set(' Pilih Jenis Pekerjaan ')
    religi_var.set(' Pilih Agama / Keyakinan ')
    fam_var.set(' Pilih Hubungan Keluarga ')
    nat_var.set(' Pilih Kewarganegaraan ')
    pas_var.set('')
    kit_var.set('')
    ayah_var.set('')
    ibu_var.set('')
    residence_var.set(' Pilih Status Domisili  ')
    fc_kk_var.set(' Pilih Ada / Tidak  ')
    ivar.set(clear)
    ivari.set(clear)
    akte_lahir_var.set('')
    akte_mati_var.set('')


# Create Digital Clock
def digital_clock(): 
   time_live = time.strftime("%H:%M:%S")
   label.config(text=time_live) 
   label.after(200, digital_clock)

text_font= ("DIGIT LCD",'18','bold')
label = Label(window, font=text_font, bg='cyan', fg='black') 
label.place(x=500,y=3)
digital_clock()

# Create Serial Number Automatically
def reg_no():
    file=openpyxl.load_workbook('Citizen_Registration_System.xlsx')
    sheet=file.active
    row=sheet.max_row
    max_row_value=sheet.cell(row=row,column=1).value
    try:
        no_var.set(max_row_value+1)
    except:
        no_var.set("1")

no_lbl = Label(window,text="No.",font='Centaur 12 bold',bg='cyan').place(x=965,y=3)
no_var = StringVar()
no_ent = Entry(window,textvariable=no_var,width=5,font=14,bg='cyan',borderwidth=0).place(x=1000,y=4)
reg_no()

# Create Entry Frame
frm = Frame(window,width=990,height=530,bd=3,relief=RIDGE)
frm.pack(side=TOP,fill=X)

# Create Label and TextBox
kklabel=Label(frm,text="No. KK",font=14).place(x=50,y=15,anchor=W)
kk_var=StringVar()
kkEntry=Entry(frm,textvariable=kk_var,width=20,bd=1,font=font1).place(x=195,y=2)

niklabel=Label(frm,text="No. NIK",font=14).place(x=600,y=15,anchor=W)
nik_var=StringVar()
nikEntry=Entry(frm,textvariable=nik_var,width=20,bd=1,font=font1).place(x=750,y=2)

namalabel=Label(frm,text="Nama Lengkap",font=14).place(x=50,y=55,anchor=W)
nama_var=StringVar()
namaEntry=Entry(frm,textvariable=nama_var,width=35,bd=1,font=font1).place(x=195,y=42)

genlabel=Label(frm,text="Jenis Kelamin",font=14).place(x=600,y=55,anchor=W)
gen_var = StringVar()
gen_var.set(' Laki-laki ')
male_radio = Radiobutton(frm, text=" Laki-laki ", variable=gen_var, value=" Laki-laki ",font=('Times','12'))
female_radio = Radiobutton(frm, text=" Perempuan ", variable=gen_var, value=" Perempuan ",font=('Times','12'))
male_radio.place(x=750,y=42)
female_radio.place(x=850,y=42)

alamatlabel=Label(frm,text="Alamat",font=14).place(x=50,y=95,anchor=W)
alamat_var=StringVar()
alamatEntry=Entry(frm,textvariable=alamat_var,width=35,bd=1,font=font1).place(x=195,y=82)

goldarlabel=Label(frm,text="Golongan Darah",font=14).place(x=600,y=95,anchor=W)
goldar_var=StringVar()
goldar_combo=Combobox(frm,textvariable=goldar_var,values=['Belum Perikasa','A','A -','A +','B','B -','B +','AB','AB -','AB +','O','O -','O +'],width=23,state='r',font=('Times','12'))
goldar_var.set(" Pilih Golongan Darah ")
goldar_combo.place(x=750,y=82)

kotalabel=Label(frm,text="Kota Kelahiran",font=14).place(x=50,y=135,anchor=W)
kota_var=StringVar()
kotaEntry=Entry(frm,textvariable=kota_var,width=35,bd=1,font=font1).place(x=195,y=122)

birth_date_label=Label(frm, text="Tanggal Lahir",font=14).place(x=600,y=135,anchor=W)
birthdatestring = '03-12-1974'
birth_date_var = datetime.datetime.strptime(birthdatestring,'%d-%m-%Y')
birth_date_entry=DateEntry(frm,textvariable=birth_date_var,date_pattern="dd - mm - yyyy",font=('Times','12'),width=23)
birth_date_entry.place(x=750,y=122)

martial_status_label=Label(frm,text="Status Perkawinan",font=14).place(x=50,y=175,anchor=W)
martial_var = StringVar()
martial_status_combo = Combobox(frm, textvariable=martial_var,values=[" Belum / Tidak Kawin "," Kawin "," Cerai Hidup"," Cerai Mati"],width=23,state='r',font=('Times','12'))
martial_var.set(" Pilih Status Perkawinan ")
martial_status_combo.place(x=195,y=162)

martial_date_label=Label(frm, text="Tanggal Perkawinan",font=14).place(x=600,y=175,anchor=W)
martialdatestring = '29-10-1968'
martial_date_var = datetime.datetime.strptime(martialdatestring,'%d-%m-%Y')
martial_date_entry=DateEntry(frm,textvariable=martial_date_var,date_pattern="dd - mm - yyyy",font=('Times','12'),width=23)
martial_date_entry.place(x=750,y=162)

educ_label=Label(frm,text=" Pendidikan Akhir ",font=14).place(x=50,y=215,anchor=W)
educ_var=StringVar()
educ_combo = Combobox(frm,textvariable=educ_var,values=['Belum/Tidak Sekolah','PAUD','TKK','Sekolah Dasar','SLTP/SMP','SLTA/SMA/SMK','Mahasiswa','Sarjana Diploma - 1','Sarjana Diploma - 2','Sarjana Diploma - 3','Sarjana Diploma - 4','Sarjana Strata - 1','Sarjana Strata - 2','Sarjana Strata - 3'],width=23,state='r',font=('Times','12'))
educ_var.set(" Pilih Pendidikan ")
educ_combo.place(x=195,y=202)

job_label=Label(frm,text="Jenis Pekerjaan",font=14).place(x=600,y=215,anchor=W)
job_var=StringVar()
job_combo = Combobox(frm,textvariable=job_var,values=['Belum/Tidak Bekerja','Pelajar/Mahasiswa','Asisten Rumah Tangga','ASN / PNS','Anggota TNI','Anggota POLRI','Pedagang','Wirausaha UMKM','Karyawan Swasta','Karyawan BUMD','Karyawan BUMN','GURU','DOSEN','Jurnalis/Wartawan','Notaris','Pengacara','ADVOKAD','Paramedis/Perawat','Apoteker','Dokter Umum','Dokter Spesialis','Dokter Hewan','Bidan','Pekerja Konstruksi','Tukang Batu','Tukang Kayu','Tukang Listrik','Tukang Pipa Air','Tukang Las','Penata Rambut / Barber','Penata Rias / Salon','Penata Busana / Tailor','Pensiunan ASN / PNS','Pensiunan Swasta','Pensiunan BUMD','Pensiunan BUMN','Uztad / Mubaliq','Imam Masjid','MODIN'],width=23,state='r',font=('Times','12'))
job_var.set(" Pilih Jenis Pekerjaan ")
job_combo.place(x=750,y=202)

religilabel=Label(frm,text="Agama",font=14).place(x=50,y=255,anchor=W)
religi_var=StringVar()
religicombo=Combobox(frm,textvariable=religi_var,values=['Islam','Kristen','Katolik','Hindu','Budha','Kong Hu Chu','Aliran Kepercayaan'],width=23,state='r',font=('Times','12'))
religi_var.set(" Pilih Agama / Keyakinan ")
religicombo.place(x=195,y=242)

famlabel=Label(frm,text="Hubungan Keluarga",font=14).place(x=600,y=255,anchor=W)
fam_var=StringVar()
famcombo=Combobox(frm,textvariable=fam_var,values=['Kepala Keluarga','Istri','Anak','Cucu','Menantu'],width=23,state='r',font=('Times','12'))
fam_var.set(" Pilih Hubungan Keluarga ")
famcombo.place(x=750,y=242)

natlabel=Label(frm,text="Kewarganegaraan",font=14).place(x=50,y=295,anchor=W)
nat_var=StringVar()
natcombo=Combobox(frm,textvariable=nat_var,values=['WNI (Indonesia)','WNA (Asing)'],width=23,state='r',font=('Times','12'))
nat_var.set(" Pilih Kewarganegaraan ")
natcombo.place(x=195,y=282)

paslabel=Label(frm,text="No. Pasport",font=14).place(x=50,y=335,anchor=W)
pas_var=StringVar()
pasentry=Entry(frm,textvariable=pas_var,width=20,bd=1,font=font1).place(x=195,y=322)

kitlabel=Label(frm,text="No. K I T A P",font=14).place(x=600,y=335,anchor=W)
kit_var=StringVar()
kitentry=Entry(frm,textvariable=kit_var,width=20,bd=1,font=font1).place(x=750,y=322)

ayahlabel=Label(frm,text="Nama Ayah",font=14).place(x=50,y=375,anchor=W)
ayah_var=StringVar()
ayahentry=Entry(frm,textvariable=ayah_var,width=20,bd=1,font=font1).place(x=195,y=362)

ibulabel=Label(frm,text="Nama Ibu",font=14).place(x=600,y=375,anchor=W)
ibu_var=StringVar()
ibuentry=Entry(frm,textvariable=ibu_var,width=20,bd=1,font=font1).place(x=750,y=362)

# Create Horizontal Separator =======================================================
s = ttk.Style()
s.configure('grey.TSeparator', background='grey')
sep = ttk.Separator(frm,orient=HORIZONTAL,style='grey.TSeparator')
sep.place(x=0,y=395,relwidth=2,height=10)

# Create Checkbox
residence_status_label=Label(frm,text="Status Domisili",font=14).place(x=50,y=420,anchor=W)
residence_var = StringVar()
residence_status_menu = OptionMenu(frm, residence_var, " Menetap "," Tidak Tinggal "," Sewa / Kontrak "," Pindah "," Meninggal ")
residence_var.set(" Pilih Status Domisili  ")
residence_status_menu.place(x=193,y=405)

fc_kk_label=Label(frm,text="Fotocopy KK",font=14).place(x=600,y=420,anchor=W)
fc_kk_var = StringVar()
fc_kk_menu = OptionMenu(frm,fc_kk_var, " Ada "," Tidak ada ")
fc_kk_var.set(" Pilih Ada / Tidak  ")
fc_kk_menu.place(x=750,y=405)

def show():
    lbl=Label(frm,text=ivar.get()).place(x=216,y=449)   
ivar = StringVar()
akte_lahir_label = Label(frm,text="Akte Kelahiran:",font=14)
akte_lahir_check = Checkbutton(frm,text="Tidak ada",variable=ivar,onvalue="    Ada     ",offvalue="Tidak ada",command=show)
akte_lahir_check.deselect()
akte_lahir_label.place(x=50,y=448)
akte_lahir_check.place(x=193,y=448)

akte_lahir_lbl=Label(frm,text="No. Akte Kelahiran",font=14).place(x=50,y=500,anchor=W)
akte_lahir_var=StringVar()
akte_lahir_ent=Entry(frm,textvariable=akte_lahir_var,width=20,bd=1,font=font1).place(x=193,y=486)

def show1():
    lbl=Label(frm,text=ivari.get()).place(x=774,y=449)
ivari = StringVar()
akte_mati_label = Label(frm,text="Akte Kelahiran:",font=14)
akte_mati_check = Checkbutton(frm,text="Tidak ada",variable=ivari,onvalue="    Ada     ",offvalue="Tidak ada",command=show1)
akte_mati_check.deselect()
akte_mati_label.place(x=600,y=448)
akte_mati_check.place(x=750,y=448)

akte_mati_lbl=Label(frm,text="No. Akte Kematian",font=14).place(x=600,y=500,anchor=W)
akte_mati_var=StringVar()
akte_mati_ent=Entry(frm,textvariable=akte_mati_var,width=20,bd=1,font=font1).place(x=750,y=486)

# Create define search Data 
def search():
# Load the workbook
    file = openpyxl.load_workbook('Citizen_Registration_System.xlsx')
    sheet = file.active
# Define the headers
    headers = [
    'ser_num',
    'kk',
    'nik',
    'nama',
    'gen',
    'alamat',
    'goldar',
    'kota',
    'birth_date',
    'martial',
    'martial_date',
    'educ',
    'job',
    'religi',
    'fam',
    'nat',
    'pas',
    'kit',
    'ayah',
    'ibu',
    'residence',
    'fc_kk',
    'akte_lahir',
    'akte_mati',
    'no_akte_lahir',
    'no_akte_mati'
]

# Define the search query
    query = {
    'ser_num'=no_var.get(),
    kk=kk_var.get(),
    nik=nik_var.get(),
    nama=nama_var.get(),
    gen=gen_var.get(),
    alamat=alamat_var.get(),
    goldar=goldar_var.get(),
    kota=kota_var.get(),
    birth_date=birth_date_entry.get(),
    martial=martial_var.get(),
    martial_date=martial_date_entry.get(),
    educ=educ_var.get(),
    job=job_var.get(),
    religi=religi_var.get(),
    fam=fam_var.get(),
    nat=nat_var.get(),
    pas=pas_var.get(),
    kit=kit_var.get(),
    ayah=ayah_var.get(),
    ibu=ibu_var.get(),
    residence=residence_var.get(),
    fc_kk=fc_kk_var.get(),
    akte_lahir=akte_lahir_check.get(),
    akte_mati=akte_mati_check.get(),
    no_akte_lahir=akte_lahir_var.get(),
    no_akte_mati=akte_mati_var.get()
}
# Search for the data
for row in sheet.iter_rows(min_row=2, values_only=True):
    if all([row[i] == query[header] for i, header in enumerate(headers)]):
        print(row)

# Create Save Button
save_button=Button(window,text='S A V E',width=15,height=1,compound=LEFT,font='David 12 bold',bg='lightgreen',fg='black',command=submit)
save_button.place(x=10,y=565)

# Create Edit Button
edit_button=Button(window,text='E D I T',width=15,height=1,compound=LEFT,font='David 12 bold',bg='orange',fg='black',command=edit)
edit_button.place(x=180,y=565)

# Create Search Entry
search_var=StringVar()
search_ent=Entry(window,textvariable=search_var,width=18,bd=1,font='Constantia 14 bold')
search_ent.place(x=512,y=565)
# Create Search Button
search_button=Button(window,text='S E A R C H NIK >>',width=15,height=1,compound=LEFT,font='David 12 bold',bg='yellow',fg='black',command=search)
search_button.place(x=350,y=565)

# Create Cancel Button
delete_button=Button(window,text='D E L E T E',width=15,height=1,compound=LEFT,font='David 12 bold',bg='pink',fg='black',command=clear)
delete_button.place(x=723,y=565)

# Create Exit Button
exit_button=Button(window,text='E X I T',width=15,height=1,compound=LEFT,font='David 12 bold',bg='white',fg='red',command=quit)
exit_button.place(x=892,y=565)

# Create Buttom Frame
lbl_buttom = Label(window,text=" Email : paulwincomptech@gmail.com ",font=("Times","14","bold"),bg='cyan',relief=RIDGE,anchor=E)
lbl_buttom.pack(side=BOTTOM,fill=X)

# Loop to window
mainloop()
