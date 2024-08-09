import os
import shutil
import sys
import time
from tkinter import *
from tkinter import filedialog
from customtkinter import *
from CustomTkinterMessagebox import CTkMessagebox
from os import walk
from win32api import *
from datetime import datetime
from datetime import date
import time
import csv
import json
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from odf.opendocument import OpenDocumentText
from odf.style import Style, TextProperties, ParagraphProperties, TabStop, TabStops
from odf.text import H, P, Span
from odf import teletype
from email.mime.application import MIMEApplication

odt_file = None
path = None
csv_file = None
folder_path = None
csv_list = []
checkbox_paths = {}
bug_dict = {}
files_list = []
bugs_list_01 = []

choiced_data = []


all_programs = [
    "CAD",
    "CAD Licencja",
    "Dystrybutor baz",
    "Instalator",
    "Konwerter mdb",
    "Marketing",
    "Panel aktualizacji",
    "Podesty",
    "Tłumaczenia",
    "aktualizator internetowy",
    "analytics",
    "asystent pobierania",
    "bazy danych",
    "deinstalator",
    "dokumentacja",
    "dot4cad",
    "drzwi i okna",
    "edytor bazy płytek",
    "edytor szafek bazy kuchennej",
    "edytor szafek użytkownika",
    "edytor ścian",
    "elementy dowolne",
    "export3D",
    "instalator baz danych",
    "instalator programu",
    "konwerter",
    "kreator ścian",
    "launcher",
    "listwy",
    "manager",
    "obrót 3d / przesuń",
    "obserVeR",
    "przeglądarka PDF",
    "render",
    "szafy wnękowe",
    "ukrywacz",
    "wersja Leroy Merlin",
    "wersja Obi",
    "wizualizacja",
    "wstawianie elementów wnętrzarskich",
    "wstawianie elementów wnętrzarskich w wizualizacji",
    "zestawienie płytek, farb, fug, klejów",
    "CAD Rozkrój"
]

kitchen_pro = [
    "blaty",
    "wstawianie elementów agd",
    "wstawianie szafek kuchennych",
    "wycena",
]



def load_data():
    global path
    data_file = f"{get_script_path()}\\paths\\path.txt"
    try:
        with open(data_file, "r") as f:
            for line in f:
                path = line
    except:
        FileNotFoundError

def get_version_number(file_path):
    try:
        File_information = GetFileVersionInfo(file_path, "\\")
        ms_file_version = File_information["FileVersionMS"]
        ls_file_version = File_information["FileVersionLS"]

        list_of_versions = [
            str(HIWORD(ms_file_version)),
            str(LOWORD(ms_file_version)),
            str(HIWORD(ls_file_version)),
            str(LOWORD(ls_file_version)),
        ]
        return ".".join(list_of_versions)
    except:
        return " "

def get_creation_date(path, long=True):
    time_created = time.ctime(os.path.getmtime(path))
    t_obj = time.strptime(time_created)
    if long == True:
        return time.strftime("%Y-%m-%d %H:%M:%S", t_obj)
    elif long == False:
        return time.strftime("%d.%m.%Y", t_obj)

def make_path_list(folder_p):
    pathName_dictionary = {}
    for dirpath, dirnames, filenames in walk(folder_p):
        for file_name in filenames:
            path = os.path.abspath(os.path.join(dirpath, file_name))
            pathName_dictionary[path] = file_name
    return pathName_dictionary

def sort_files_del_from_dict(file_list, bug_d):
    list_of_names = []
    list_without_duplicates = []

    for i in file_list:
        list_of_names.append(os.path.basename(i))

    [list_without_duplicates.append(x) for x in list_of_names if x not in list_without_duplicates]

    for n in range(len(list_without_duplicates)):
        condition = lambda x: os.path.basename(x) == list_without_duplicates[n]

        filtered_list = [x for x in file_list if condition(x)]
        time_list = []

        for item in filtered_list:
            time_list.append(os.path.getmtime(item))
            x = time_list.index(max(time_list))

        del filtered_list[x]

        for a in filtered_list:
            del bug_d[a]

    listed_dict = [x for x in bug_d.keys()]

    for x in listed_dict:
        if x.endswith(".txt"):
            del bug_d[x]
    return bug_d

def make_list_to_cut(dict):
    indexesToCut = []
    v_list = [v for v in dict.values()]

    for i, v in enumerate(dict.values()):
        if v_list.count(v) > 1 and i not in indexesToCut:
            indexesToCut.append(i)
    return indexesToCut

def list_paths(i_list, paths):

    cuted_pathList = []
    k_list = [k for k in paths.keys()]
    for index in i_list:
        pathList = k_list
        cuted_pathList += [pathList[index]]
    return cuted_pathList

def add_lines_to_lists(data_file, all_list, rest_list):
    cat_list_with_duplicates = []
    cat_list_with_duplicates02 = []
    cat_list_LM_with_duplicates = []
    cat_list_OBI_with_duplicates = []
    cat_list_rozkroj_with_duplicates = []
    cat_list_OBI = []
    cat_list_LM = []
    cat_list02 = []
    cat_list = []
    cat_list_rozkroj = []

    for line in data_file:
        if line[3].lower() == "wersja obi":
            cat_list_OBI_with_duplicates.append(line[3])
            cat_list_OBI = []
            [
                cat_list_OBI.append(x)
                for x in cat_list_OBI_with_duplicates
                if x not in cat_list_OBI
            ]

        elif line[3].lower() == "wersja leroy merlin":
            cat_list_LM_with_duplicates.append(line[3])
            cat_list_LM = []
            [
                cat_list_LM.append(x)
                for x in cat_list_LM_with_duplicates
                if x not in cat_list_LM
            ]

        elif line[3].lower() == "cad rozkrój":
            cat_list_rozkroj_with_duplicates.append(line[3])
            cat_list_rozkroj = []
            [
                cat_list_rozkroj.append(x)
                for x in cat_list_rozkroj_with_duplicates
                if x not in cat_list_rozkroj
            ]

        elif line[3].lower() != "projekt" and line[3] in all_list:
            cat_list_with_duplicates.append(line[3])
            cat_list = []
            [cat_list.append(x) for x in cat_list_with_duplicates if x not in cat_list]

        elif line[3] in rest_list:
            cat_list_with_duplicates02.append(line[3])
            cat_list02 = []
            [
                cat_list02.append(x)
                for x in cat_list_with_duplicates02
                if x not in cat_list02
            ]

    return cat_list, cat_list02, cat_list_LM, cat_list_OBI, cat_list_rozkroj

def make_lines(m_name, data_dict, document):
    for key, value in data_dict:
        if key == m_name:
            headline = H(outlinelevel=1, stylename=heading03_style, text=key.upper())
            make_add_paragraph("", document)
            document.text.addElement(headline)

            for line in value:
                line_ = P(stylename=tabparagraphstyle01)
                boldpart = Span(stylename=boldstyle, text=line[0])
                line_.addElement(boldpart)
                text = f"\t{line[2]}"
                teletype.addTextToElement(line_, text)
                document.text.addElement(line_)
             
def make_bug_dict(file, m_name):
    list_00 = []
    for line in file:
        if line[3] == m_name:
            var_01 = (line[0][3:], " ", line[1])
            list_00.append(var_01)
    bug_dict.update({m_name: list_00})
    return bug_dict

def save_with_current_day(document):
    doc_name = datetime.today().strftime("%d.%m.%Y")
    file = os.path.join(get_script_path(), doc_name + ".odt")
    document.save(file)

def write_changed_files(dict_of_files):
    list_of_lines_to_write = []
    for key, value in dict_of_files.items():
        file_creation_date = get_creation_date(key, long=False)
        file_version = get_version_number(key)
        file_name = value
        list_of_lines_to_write.append(
            f"{file_name} {file_version}\t{file_creation_date}"
        )
    return list_of_lines_to_write

def make_add_paragraph(text, document):
    var_01 = P(stylename=paragraph_style00, text=text)
    document.text.addElement(var_01)

def make_add_heading(text, document):
    var_01 = H(outlinelevel=1, stylename=heading01_style, text=text)
    document.text.addElement(var_01)

def get_and_display_path():
    choice = var_0.get()
    global folder_path

    if choice == 1:
        for n in paths_dir:
            if n.endswith("hotfix") == True and n.endswith("_op") == False:
                hotfix_cat = n
                hotfix_path = os.path.join(path + "\\" + hotfix_cat)
        label1.configure(text=hotfix_cat)
        folder_path = hotfix_path
        button3.configure(state="normal")
        button_c.configure(state="normal")
        button2.configure(state="normal")

    elif choice == 2:
        try:
            for n in paths_dir:
                if (
                    "feat" not in n
                    and n.endswith("_op") == False
                    and n.endswith("hotfix") == False
                    and n.endswith('.odt') == False
                    and n.endswith('.txt') == False
                    and n != "_dok"
                ):
                    new_version_cat = n
                    new_version_path = os.path.join(path + "\\" + new_version_cat)

            label1.configure(text=new_version_cat)
            folder_path = new_version_path
            button3.configure(state="normal")
            button_c.configure(state="normal")

        except:
            UnboundLocalError
            folder_path = ""
            label1.configure(text="Brak paczki z nową wersją")
            button3.configure(state="disabled")
            button2.configure(state="disabled")

def get_script_path():
    return os.path.dirname(os.path.realpath(sys.argv[0]))

def ask_for_dir(value):
    global folder_path
    global csv_file
    if value == 1:
        csv_dir = filedialog.askopenfilename()
        if csv_dir.endswith(".csv"):
            s_dir = csv_dir.split("/")
            label0.configure(text=s_dir[-1])
            csv_file = csv_dir
            button_csv.configure(state='normal')

        else:
            csv_file = ""
            label0.configure(text="Wybierz prawidłowy plik")

    elif value == 2:
        pack_dir = filedialog.askdirectory()
        s_pack_dir = pack_dir.split("/")

        label1.configure(text=s_pack_dir[-1])
        folder_path = pack_dir
        button3.configure(state="normal")
        button2.configure(state="normal")
        button_c.configure(state="normal")
        

def make_list(folder, csv_f):

    doc = OpenDocumentText()
    doc_s = doc.styles

    heading01_style = Style(name="Heading 1", family="paragraph")
    heading01_style.addElement(
        TextProperties(
            attributes={"fontsize": "16pt", "fontweight": "bold", "fontfamily": "Calibri"}
        )
    )
    heading01_style.addElement(ParagraphProperties(attributes={"textalign": "center"}))
    doc_s.addElement(heading01_style)

    heading02_style = Style(name="Heading 2", family="paragraph")
    heading02_style.addElement(
        TextProperties(
            attributes={"fontsize": "14pt", "fontweight": "bold", "fontfamily": "Calibri"}
        )
    )
    heading02_style.addElement(ParagraphProperties(attributes={"textalign": "center"}))
    doc_s.addElement(heading02_style)

    heading03_style = Style(name="Heading 3", family="paragraph")
    heading03_style.addElement(
        TextProperties(
            attributes={"fontsize": "11pt", "fontweight": "bold", "fontfamily": "Calibri"}
        )
    )
    heading03_style.addElement(ParagraphProperties(lineheight="145%"))
    doc_s.addElement(heading03_style)


    underline_style = Style(name="Underline", family="text")
    u_prop = TextProperties(
        attributes={
            "textunderlinestyle": "solid",
            "textunderlinewidth": "auto",
            "textunderlinecolor": "font-color",
        }
    )

    underline_style.addElement(u_prop)
    doc_s.addElement(underline_style)

    boldstyle = Style(name="Bold", family="text")
    boldstyle.addElement(TextProperties(attributes={"fontweight": "bold"}))
    doc_s.addElement(boldstyle)

    paragraph_style00 = Style(
        name="paragraph",
        family="paragraph",
    )
    paragraph_style00.addElement(
        TextProperties(attributes={"fontsize": "11pt", "fontfamily": "Calibri"})
    )
    paragraph_style00.addElement(ParagraphProperties(lineheight="135%"))
    doc_s.addElement(paragraph_style00)

    tabstops_style = TabStops()
    tabstop_style = TabStop(position="5cm")
    tabstops_style.addElement(tabstop_style)
    tabstoppar = ParagraphProperties()
    tabstoppar.addElement(tabstops_style)
    tabparagraphstyle = Style(name="Question", family="paragraph")
    tabparagraphstyle.addElement(
        TextProperties(attributes={"fontsize": "11pt", "fontfamily": "Calibri"})
    )
    tabparagraphstyle.addElement(ParagraphProperties(lineheight="135%")) 
    tabparagraphstyle.addElement(tabstoppar)
    doc_s.addElement(tabparagraphstyle)

    tabstop_style01 = TabStop(position="1.2cm")
    tabstops_style.addElement(tabstop_style01)
    tabstoppar01 = ParagraphProperties()
    tabstoppar01.addElement(tabstops_style)
    tabparagraphstyle01 = Style(name="Question", family="paragraph")
    tabparagraphstyle01.addElement(
            TextProperties(attributes={"fontsize": "11pt", "fontfamily": "Calibri"})
        )
    tabparagraphstyle01.addElement(ParagraphProperties(lineheight="135%"))
    tabparagraphstyle01.addElement(tabstoppar01)
    doc_s.addElement(tabparagraphstyle01)


    if len(checkbox_paths) == 0:
        all_paths = make_path_list(folder)
        indexesList = make_list_to_cut(all_paths)
        cuted_paths = list_paths(indexesList, all_paths)
        target_paths = sort_files_del_from_dict(cuted_paths, all_paths)
    else:
        target_paths = checkbox_paths


    if len(choiced_data) == 0:
        with open(csv_f, mode="r", encoding="utf-8") as file:
            csv_file = csv.reader(file)
            data = [tuple(row) for row in csv_file]
    else:
        data = choiced_data


    cat_list, cat_list02, cat_list_LM, cat_list_OBI, cat_list_rozkroj = add_lines_to_lists(
        data, all_programs, kitchen_pro
    )

    make_add_heading(
        f"Aktualizacja z dnia {datetime.today().strftime('%d.%m.%Y')} ", doc
    )
    make_add_paragraph("", doc)

    if len(cat_list) > 0:
        heading_0 = H(outlinelevel=1, stylename=heading02_style, text="")
        underlinedpart = Span(
            stylename=underline_style,
            text="Zmiany wspólne dla programów CAD Decor PRO i CAD Decor oraz CAD Kuchnie",
        )
        heading_0.addElement(underlinedpart)
        doc.text.addElement(heading_0)
        make_add_paragraph("", doc)

        for category in cat_list:
            make_bug_dict(data, category)
            make_lines(category, bug_dict.items(), doc)

    if len(cat_list02) > 0:
        make_add_paragraph("", doc)
        heading_1 = H(outlinelevel=1, stylename=heading02_style, text="")
        underlinedpart = Span(
            stylename=underline_style,
            text="Zmiany wspólne dla programów CAD Decor PRO oraz CAD Kuchnie",
        )
        heading_1.addElement(underlinedpart)
        doc.text.addElement(heading_1)
        make_add_paragraph("", doc)

        for category in cat_list02:
            make_bug_dict(data, category)
            make_lines(category, bug_dict.items(), doc)

    if len(cat_list_LM) > 0:
        make_add_paragraph("", doc)
        heading_2 = H(outlinelevel=1, stylename=heading02_style, text="")
        underlinedpart = Span(
            stylename=underline_style, text="Zmiany dla programu w wersji LM"
        )
        heading_2.addElement(underlinedpart)
        doc.text.addElement(heading_2)
        make_add_paragraph("", doc)

        for category in cat_list_LM:
            make_bug_dict(data, category)
            make_lines(category, bug_dict.items(), doc)

    if len(cat_list_OBI) > 0:
        make_add_paragraph("", doc)
        heading_3 = H(outlinelevel=1, stylename=heading02_style, text="")
        underlinedpart = Span(
            stylename=underline_style, text="Zmiany dla programu w wersji OBI"
        )
        heading_3.addElement(underlinedpart)
        doc.text.addElement(heading_3)
        make_add_paragraph("", doc)

        for category in cat_list_OBI:
            make_bug_dict(data, category)
            make_lines(category, bug_dict.items(), doc)


    if len(cat_list_rozkroj) > 0:
        make_add_paragraph("", doc)
        heading_4 = H(outlinelevel=1, stylename=heading02_style, text="")
        underlinedpart=Span(stylename=underline_style, text="Zmiany dla programu CAD Rozkrój")
        heading_4.addElement(underlinedpart)
        doc.text.addElement(heading_4)
        make_add_paragraph("", doc)

        for category in cat_list_rozkroj:
            make_bug_dict(data, category)
            make_lines(category, bug_dict.items(), doc)
    

    make_add_paragraph("", doc)
    make_add_paragraph("ZMIENIONE PLIKI:", doc)

    list_of_lines = write_changed_files(target_paths)
    sorted_list_of_lines = sorted(list_of_lines, key=str.casefold)

    for line in sorted_list_of_lines:
        tabp = P(stylename=tabparagraphstyle)
        teletype.addTextToElement(tabp, line)
        doc.text.addElement(tabp)

    save_with_current_day(doc)

def copy_pack(folder): 
    global a_dirs_to_delete
    a_dirs_to_delete = []
    global new_dir

    try:
        new_dir = get_script_path() + "/" + datetime.today().strftime("%Y-%m-%d") + " x64"
        shutil.copytree(folder, new_dir)
    except: 
        FileExistsError
        new_dir = get_script_path() + "/" + datetime.today().strftime("%Y-%m-%d") + " x64(1)"
        shutil.copytree(folder, new_dir)

    directory_list =  make_path_list(new_dir)

  
    for i in directory_list:
        if i.endswith(".txt"):
            os.remove(i)
       
    i_list = make_list_to_cut(directory_list)


    if len(i_list) > 0:
        cuted_p = list_paths(i_list, directory_list)
       
        t_paths = sort_files_del_from_dict(cuted_p, directory_list)

        del_files_and_dirs(new_dir, t_paths)
        
    move_odt_file(new_dir)
    
    button4.configure(state="normal", hover=True)
    
def del_files_and_dirs(dir, paths):

    for p in make_path_list(dir):

        dir_to_remove = []
        dir_to_remove_02 = []

        if p not in paths:
            os.remove(p)
            if p.endswith('.txt') == False:
                dir_to_remove.append(os.path.dirname(p))
           
        for i in dir_to_remove:
            dir_to_remove_02.append(os.path.dirname(i))
            for i in dir_to_remove_02:
                try: 
                    os.rmdir(i)
                except: 
                    OSError
            
        for p in dir_to_remove_02:     
            if p not in a_dirs_to_delete:
                a_dirs_to_delete.append(p)

    for i in a_dirs_to_delete: 
        if i.endswith('V4_I10x64') == False:   
            shutil.rmtree(i) 
       
                

    list_from_paths = [x for x in paths.keys()]

    if files_list != 0:
        for file in files_list:
            for key in list_from_paths:
                if os.path.basename(file) == os.path.basename(key):
                    os.remove(key)
                    try:
                        os.rmdir(os.path.dirname(key))
                    except:
                        OSError

                

def move_odt_file(new_dir):
    dir = os.listdir(get_script_path())
    odt_list = []
    try:
        for file in dir:
            if file.endswith(".odt") == True:
                file_p = get_script_path() + "\\" + file
                odt_list.append(file_p)
                odt_list.sort(key=os.path.getmtime, reverse=True)
                shutil.move(odt_list[0], new_dir)
    except:
        print("error")

def move_csv_files():
    file_p = ''
    csv_files_list = []
    data_file = f"{get_script_path()}\\paths\\csv_path.txt"
    csv_list02 = []
    try:
        with open(data_file, "r") as f:
            for line in f:
                file_p = line
    except:
        FileNotFoundError
    
    try:
        csv_files_list = os.listdir(file_p)
    except:
        FileNotFoundError


    if len(csv_files_list) > 0:
        try:
            for item in csv_files_list:
                if item.endswith(".csv") == True:
                    file = file_p + "\\" + item
                    csv_list02.append(file)
                    csv_list02.sort(key=os.path.getmtime, reverse=True)
            shutil.move(csv_list02[0], get_script_path())
        except:
            print("error in move csv")
    
    get_default_csv()

def get_default_csv():
    global csv_file
    list = os.listdir(get_script_path())
    try:
        paths_dir = os.listdir(path)
    except:
        FileNotFoundError
        print("error in get default csv")
    try:
        for file in list:
            if file.endswith(".csv") == True:
                file_p = get_script_path() + "\\" + file
                csv_list.append(file_p)
                csv_list.sort(key=os.path.getmtime, reverse=True)
                splited_csv = csv_list[0].split("\\")
        label0.configure(text=splited_csv[-1]) 
        csv_file = csv_list[0]
    except:
        label0.configure(text="Wybierz plik csv")
    return csv_file, paths_dir

def send_email(user, email_receiver, kafle):

    f = open(get_script_path() + "\\" + "paths\\users.json", encoding='utf-8')
    data = json.load(f)

    if user == 1:
        user = data["users"][1]["displayname"]
    elif user == 2:
        user = data["users"][0]["displayname"]

    if email_receiver == 1:
        # emails = data["receivers"][0]["emails"]
        emails = ["tomasz.czerwinski@cadprojekt.com.pl"]
    elif email_receiver == 2:
        # emails = data["receivers"][1]["emails"]
        emails = ["tomasz.czerwinski@cadprojekt.com.pl"]


    sender_email = "tomasz.czerwinski@cadprojekt.com.pl"
    # sender_email = "testy@cadprojekt.com.pl"
    # password = "testy password"
    password = data["users"][0]["password"]

    message = MIMEMultipart("multipart")
    # message = MIMEMultipart("alternative")
    message["Subject"] = f"Aktualizacja z dnia {date.today().strftime("%d.%m.%Y")}"
    message["From"] = sender_email
    message["To"] = ', '.join(emails)
    path_to_file = odt_file
    port = data["users"][0]["port"]


    if len(kafle) > 1:
        # text = data["messeges"][1]["data"] + kafle
        text = '<font face="Segoe UI">' + data["messeges"][1]["data"] + kafle + '</font>'
    else:
        # text = data["messeges"][0]["data"]
        text = '<font face="Segoe UI">' + data["messeges"][0]["data"] + '</font>'


    footer_tomasz = """\

    <html>
        <head>
            <meta name="viewport" content="width=device-width,initial-scale=1,user-scalable=no">
        </head>
        <body style="margin: 0; padding: 20px;">
        <div style="font: normal normal normal 12px/14px Helvetica; color: #000; padding: 5px 0;">Pozdrawiam / Kind regards</div>
            <div style="padding: 10px 0;">
                <div style="display: flex; flex-wrap: wrap; align-items: center; margin: 0 -20px;">
                    <div style="width: 340px;">
                        <div style="padding: 0 20px">						
                            <div style="font: normal normal bold 20px/24px Helvetica; color: #000; padding: 5px 0;">Tomasz Czerwiński</div>
                            <div style="font: normal normal normal 14px/16px Helvetica; color: #000;">Junior Software Tester</span></div>
                        </div>
                    </div>
                    <div style="width: 330px;display: flex;align-items: center;align-self: stretch;height: 56px;border-left: 1px solid #000;margin-left: -1px;">
                        <div style="padding-left: 20px;">
                                <a style="display: block; text-decoration: none; font: normal normal normal 11px/12px Helvetica; color: #000; padding-top: 5px;" href="mailto:tomasz.czerwinski@cadprojekt.com.pl">tomasz.czerwinski@cadprojekt.com.pl</a>
                            <a style="display: block; text-decoration: none; font: normal normal normal 11px/12px Helvetica; color: #000; padding-top: 5px;" href="mailto:testy@cadprojekt.com.pl">testy@cadprojekt.com.pl</a>
                        </div>
                    </div>
                    <div style="padding: 0 20px; align-self: flex-end;">
                        <img style="height: 24px; width: auto; display: block; margin-top: 10px;" src="https://cadprojekt.com.pl/zasoby/email/logo_spjawna.png" alt="CAD PROJEKT">
                    </div>
                </div>
            </div>
            <div style="font: normal normal normal 10px/14px Helvetica; color: #666; padding: 10px 0; width: 815px; max-width: 100%; border-top: 3px solid #F3951A;">
                CAD Projekt K&A Sp.j. Dąbrowski, Sterczała, Sławek  | Poznański Park Naukowo-Technologiczny <br>
                ul. Rubież 46 | 61-612 Poznań | Poland | NIP 779-00-34-266 | REGON 632223660 | <a style="text-decoration: none; color: #666" href="https://www.cadprojekt.com.pl">www.cadprojekt.com.pl</a> 
            </div>
        
            <a style="display: block; font: normal normal normal 10px/14px Helvetica; color: #666; padding: 10px 0; text-decoration: none;" href="https://cadprojekt.com.pl/polityka-prywatnosci-sp-j" target="_blank">
                Zapoznaj się z zasadami przetwarzania Twoich danych osobowych.
            </a>
        </body>
    </html>

    """

    footer_kinga = """  

<html>
    <head>
        <meta name="viewport" content="width=device-width,initial-scale=1,user-scalable=no">
    </head>
    <body style="margin: 0; padding: 20px;">
	<div style="font: normal normal normal 12px/14px Helvetica; color: #000; padding: 5px 0;">Pozdrawiam / Kind regards</div>
        <div style="padding: 10px 0;">
            <div style="display: flex; flex-wrap: wrap; align-items: center; margin: 0 -20px;">
                <div style="width: 340px;">
                    <div style="padding: 0 20px">						
                        <div style="font: normal normal bold 20px/24px Helvetica; color: #000; padding: 5px 0;">Kinga Kaczanowska</div>
                        <div style="font: normal normal normal 14px/16px Helvetica; color: #000;">Junior Software Tester</span></div>
                    </div>
                </div>
                <div style="width: 330px;display: flex;align-items: center;align-self: stretch;height: 56px;border-left: 1px solid #000;margin-left: -1px;">
                    <div style="padding-left: 20px;">
                             <a style="display: block; text-decoration: none; font: normal normal normal 11px/12px Helvetica; color: #000; padding-top: 5px;" href="mailto:kinga.kaczanowska@cadprojekt.com.pl">kinga.kaczanowska@cadprojekt.com.pl</a>
                        <a style="display: block; text-decoration: none; font: normal normal normal 11px/12px Helvetica; color: #000; padding-top: 5px;" href="mailto:testy@cadprojekt.com.pl">testy@cadprojekt.com.pl</a>
                    </div>
                </div>
                <div style="padding: 0 20px; align-self: flex-end;">
                    <img style="height: 24px; width: auto; display: block; margin-top: 10px;" src="https://cadprojekt.com.pl/zasoby/email/logo_spjawna.png" alt="CAD PROJEKT">
                </div>
            </div>
        </div>
        <div style="font: normal normal normal 10px/14px Helvetica; color: #666; padding: 10px 0; width: 815px; max-width: 100%; border-top: 3px solid #F3951A;">
            CAD Projekt K&A Sp.j. Dąbrowski, Sterczała, Sławek  | Poznański Park Naukowo-Technologiczny <br>
			ul. Rubież 46 | 61-612 Poznań | Poland | NIP 779-00-34-266 | REGON 632223660 | <a style="text-decoration: none; color: #666" href="https://www.cadprojekt.com.pl">www.cadprojekt.com.pl</a> 
        </div>
       
        <a style="display: block; font: normal normal normal 10px/14px Helvetica; color: #666; padding: 10px 0; text-decoration: none;" href="https://cadprojekt.com.pl/polityka-prywatnosci-sp-j" target="_blank">
            Zapoznaj się z zasadami przetwarzania Twoich danych osobowych.
        </a>
    </body>
</html>              

"""

    if user == data["users"][0]["displayname"]:
        footer = footer_tomasz
    elif user == data["users"][1]["displayname"]:
        footer = footer_kinga

    # Turn these into plain/html MIMEText objects
    part1 = MIMEText(text, "html")
    part2 = MIMEText(footer, "html")

    with open(path_to_file, 'rb') as file:
        message.attach(MIMEApplication(file.read(), Name=os.path.basename(path_to_file)))

    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(part1)
    message.attach(part2)

    # Create secure connection with server and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("poczta.cadprojekt.com.pl", port, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(
            sender_email, emails, message.as_string()
        )

def open_new_window(pack):
    global odt_file

    def ask_for_list_to_send():
        global odt_file
        x = filedialog.askopenfilename()
        while x.endswith('.odt') != True:
            x = filedialog.askopenfilename()
        else:
            odt_file = x
            label01.configure(text=os.path.basename(odt_file))
            
    

    button4.configure(state="disabled", hover=True)

    is_file = True
    odt_name = date.today().strftime("%d.%m.%Y") + ".odt"
  
    newWindow = CTkToplevel(gui)
    newWindow.title("email sender")
    newWindow.geometry("380x400")

    x = gui.winfo_x() + gui.winfo_width()//2 - newWindow.winfo_width()//2
    y = gui.winfo_y() + gui.winfo_height()//2 - newWindow.winfo_height()//2
    newWindow.geometry(f"+{x}+{y}")
    newWindow.after(10, newWindow.lift)
    
    var_0 = IntVar()
    var_1 = IntVar()    

    if os.path.isfile(pack + "\\" + odt_name):
        odt_file = pack + "\\" + odt_name
        label01 = CTkLabel(newWindow, text=odt_name, fg_color="transparent", font=("Consolas", 14))
        label01.pack(padx=(0,0), pady=(10,0), anchor=CENTER)
    else:
        label01 = CTkLabel(newWindow, text="Dodaj plik z listą", fg_color="transparent", font=("Consolas", 14))
        label01.pack(padx=(0,0), pady=(10,0), anchor=CENTER)
        is_file = False

    e_button01 = CTkButton(newWindow, text="Dodaj listę", width=160, height=40, font=("Consolas", 14), command=ask_for_list_to_send)
    e_button01.pack(padx=(0,0), pady=(15,20), anchor=CENTER)


    e_frame = CTkFrame(newWindow)
    e_frame.configure(border_width=2, fg_color="transparent")
    e_frame.pack()

    e_radiobutton01 = CTkRadioButton(e_frame, text="Kinga", variable=var_0, value=1, font=("Consolas", 14))
    e_radiobutton01.pack(pady=20, padx=(30, 5), side="left", anchor=N)
    if is_file == False:
        e_radiobutton01.configure(state="disabled")

    e_radiobutton02 = CTkRadioButton(e_frame, text="Tomasz", variable=var_0, value=2, font=("Consolas", 14))
    e_radiobutton02.pack(pady=20, padx=(5, 15), side="left", anchor=N)
    if is_file == False:
        e_radiobutton02.configure(state="disabled")

    e_frame02 = CTkFrame(newWindow)
    e_frame02.configure(border_width=2, fg_color="transparent")
    e_frame02.pack(pady=(20,0), padx=(0,0))

    e_radiobutton03 = CTkRadioButton(e_frame02, text="Serwis", variable=var_1, value=1, font=("Consolas", 14))
    e_radiobutton03.pack(pady=20, padx=(30, 5), side="left", anchor=N)
    if is_file == False:
        e_radiobutton03.configure(state="disabled")

    e_radiobutton04 = CTkRadioButton(e_frame02, text="Wszyscy", variable=var_1, value=2, font=("Consolas", 14))
    e_radiobutton04.pack(pady=20, padx=(5, 15), side="left", anchor=N)
    if is_file == False:
        e_radiobutton04.configure(state="disabled")

    
    x = pack + "\\" + "MainFiles\\V4_I10x64\\kafle.dll"

    try:
        kafle = get_version_number(x)
    except:
        print("error with kafle")

    if os.path.isfile(x):
        label02 = CTkLabel(newWindow, text="Wersja programu: " + kafle, fg_color="transparent", font=("Consolas", 14))     
    else:
        label02 = CTkLabel(newWindow, text="Wersja programu: " + "bez zmian.", fg_color="transparent", font=("Consolas", 14))

    label02.pack(padx=(0,0), pady=(15,0), anchor=CENTER)

    e_button02 = CTkButton(newWindow, text="Wyślij email", width=160, height=40, font=("Consolas", 14), command=lambda : send_email(var_0.get(), var_1.get(), kafle))
    e_button02.pack(padx=(0,0), pady=(20,20), anchor=CENTER)
    if is_file == False:
        e_button02.configure(state="disabled")

    newWindow.mainloop()

def checkbox_event(file, var):
    global files_list
    
    if var.get() == False:
        files_list.append(file)
    elif var.get() == True:
        files_list.remove(file)


def ok_button_function(window):

    if len(files_list) == len(checkbox_paths):
        x = CTkMessagebox.messagebox(
        title="Błąd!",
        text='Wybierz chociaż jeden plik',
        button_text="OK",
    )
    else:
        for f in files_list:
            del checkbox_paths[f]
        window.destroy()

def open_choice_window(path):
    global files_list
    files_list = []

    global checkbox_paths
    choice_window = CTkToplevel(gui)
    choice_window.title("Wybierz pliki")
   

    x = gui.winfo_x() + gui.winfo_width()//2 - choice_window.winfo_width()//2
    y = gui.winfo_y() + gui.winfo_height()//2 - choice_window.winfo_height()//2
    choice_window.geometry(f"+{x}+{y}")

    choice_window.after(10, choice_window.lift)


    all_paths = make_path_list(path)
    indexesList = make_list_to_cut(all_paths)
    cuted_paths = list_paths(indexesList, all_paths)
    checkbox_paths = sort_files_del_from_dict(cuted_paths, all_paths)


    scrollable_frame = CTkScrollableFrame(choice_window, width=230)

    scrollable_frame.pack(side="top", pady=(20, 20), padx=(20, 20))

  
    if len(checkbox_paths) > 3 and len(checkbox_paths) < 7:
        scrollable_frame.configure(height=400)
    elif len(checkbox_paths) >= 7:
        scrollable_frame.configure(height=450)


    for file in checkbox_paths:
        checkbox_var = BooleanVar(value=True)
        checkbox = CTkCheckBox(scrollable_frame, text=os.path.basename(file), font=("Consolas", 14), variable=checkbox_var, command=lambda x=file, var=checkbox_var : checkbox_event(x, var))
        checkbox.pack(pady=(20, 20), padx=(20, 20), side="top", anchor=W)

    ok_button = CTkButton(choice_window, text="Ok", width=80, height=40, font=("Consolas", 14), command=lambda x=choice_window : ok_button_function(x))
    ok_button.pack(padx=(0,0), pady=(15,20), anchor=CENTER)


def handle_bugs_list(bug, var):

    global bugs_list_01
    
    if var.get() == False:
        bugs_list_01.append(bug)
    elif var.get() == True:
        bugs_list_01.remove(bug)


def ok_button_function_bugs(window):

    if len(bugs_list_01) + 1 == len(choiced_data):
        x = CTkMessagebox.messagebox(
        title="Błąd!",
        text='Wybierz chociaż jeden plik',
        button_text="OK",
        
    )          
    elif len(bugs_list_01) + 1 != len(choiced_data):
        for i in choiced_data[:]:
            if i[0] in bugs_list_01:
                choiced_data.remove(i)
        window.destroy()


def open_bugs(file_0):

    global choiced_data
    global bugs_list_01
    bugs_list_01 = []

    list_of_csv_items = []

    bugs_window = CTkToplevel(gui)
    bugs_window.title("Wybierz bugi")
   
    x = gui.winfo_x() + gui.winfo_width()//2 - bugs_window.winfo_width()//2
    y = gui.winfo_y() + gui.winfo_height()//2 - bugs_window.winfo_height()//2
    bugs_window.geometry(f"+{x}+{y}")
    bugs_window.after(10, bugs_window.lift)

    bug_list = []

    with open(file_0, mode="r", encoding="utf-8") as item:
        file = csv.reader(item)
        list_of_csv_items = [tuple(row) for row in file]

    choiced_data = list_of_csv_items

    for a in list_of_csv_items:
        bug_list.append(a[0] + ' ' + a[1][:100] + '...')
        
    bug_list.pop(0)

    scrollable_frame_01 = CTkScrollableFrame(bugs_window, width=500)

    scrollable_frame_01.pack(side="top", pady=(20, 20), padx=(20, 20))

    if len(bug_list) > 3 and len(bug_list) < 7:
        scrollable_frame_01.configure(height=400)
    elif len(bug_list) >= 7:
        scrollable_frame_01.configure(height=450)


    for line in bug_list:
        checkbox_var = BooleanVar(value=True)
        checkbox = CTkCheckBox(scrollable_frame_01, text=line, font=("Consolas", 14), variable=checkbox_var, command=lambda x=line[:7], var=checkbox_var : handle_bugs_list(x, var))
        checkbox.pack(pady=(20, 20), padx=(20, 20), side="top", anchor=W)

    ok_button = CTkButton(bugs_window, text="Ok", width=80, height=40, font=("Consolas", 14), command=lambda x=bugs_window : ok_button_function_bugs(x))
    ok_button.pack(padx=(0,0), pady=(15,20), anchor=CENTER)



doc = OpenDocumentText()
doc_s = doc.styles

heading01_style = Style(name="Heading 1", family="paragraph")
heading01_style.addElement(
    TextProperties(
        attributes={"fontsize": "16pt", "fontweight": "bold", "fontfamily": "Calibri"}
    )
)
heading01_style.addElement(ParagraphProperties(attributes={"textalign": "center"}))
doc_s.addElement(heading01_style)

heading02_style = Style(name="Heading 2", family="paragraph")
heading02_style.addElement(
    TextProperties(
        attributes={"fontsize": "14pt", "fontweight": "bold", "fontfamily": "Calibri"}
    )
)
heading02_style.addElement(ParagraphProperties(attributes={"textalign": "center"}))
doc_s.addElement(heading02_style)

heading03_style = Style(name="Heading 3", family="paragraph")
heading03_style.addElement(
    TextProperties(
        attributes={"fontsize": "11pt", "fontweight": "bold", "fontfamily": "Calibri"}
    )
)
heading03_style.addElement(ParagraphProperties(lineheight="145%"))
doc_s.addElement(heading03_style)


underline_style = Style(name="Underline", family="text")
u_prop = TextProperties(
    attributes={
        "textunderlinestyle": "solid",
        "textunderlinewidth": "auto",
        "textunderlinecolor": "font-color",
    }
)

underline_style.addElement(u_prop)
doc_s.addElement(underline_style)

boldstyle = Style(name="Bold", family="text")
boldstyle.addElement(TextProperties(attributes={"fontweight": "bold"}))
doc_s.addElement(boldstyle)

paragraph_style00 = Style(
    name="paragraph",
    family="paragraph",
)
paragraph_style00.addElement(
    TextProperties(attributes={"fontsize": "11pt", "fontfamily": "Calibri"})
)
paragraph_style00.addElement(ParagraphProperties(lineheight="135%"))
doc_s.addElement(paragraph_style00)

tabstops_style = TabStops()
tabstop_style = TabStop(position="7cm")
tabstops_style.addElement(tabstop_style)
tabstoppar = ParagraphProperties()
tabstoppar.addElement(tabstops_style)
tabparagraphstyle = Style(name="Question", family="paragraph")
tabparagraphstyle.addElement(
        TextProperties(attributes={"fontsize": "11pt", "fontfamily": "Calibri"})
    )
tabparagraphstyle.addElement(ParagraphProperties(lineheight="135%"))
tabparagraphstyle.addElement(tabstoppar)
doc_s.addElement(tabparagraphstyle)

tabstop_style01 = TabStop(position="1.2cm")
tabstops_style.addElement(tabstop_style01)
tabstoppar01 = ParagraphProperties()
tabstoppar01.addElement(tabstops_style)
tabparagraphstyle01 = Style(name="Question", family="paragraph")
tabparagraphstyle01.addElement(
        TextProperties(attributes={"fontsize": "11pt", "fontfamily": "Calibri"})
    )
tabparagraphstyle01.addElement(ParagraphProperties(lineheight="135%"))
tabparagraphstyle01.addElement(tabstoppar01)
doc_s.addElement(tabparagraphstyle01)




gui = CTk()

gui.geometry("350x700")
gui.title("Generator")
gui.resizable(False, False)

gui.eval('tk::PlaceWindow . center')

var_0 = IntVar()

button00 = CTkButton(
    gui,
    text="*",
    width=20,
    height=8,
    font=("Consolas", 14),
    command=move_csv_files
)
button00.pack(pady=(10, 0), padx=(10,0), anchor=W)
button00.configure(border_width=1)

label0 = CTkLabel(gui, height=20, font=("Consolas", 14))
label0.pack(pady=(0,10), anchor=CENTER)


load_data()

if path == None:
    result = CTkMessagebox.messagebox(
        title="Błąd!",
        text='Dla automatycznego wykrywania paczki \n dodaj plik "path.txt" ze ścieżką.',
        button_text="OK",
    )

try:
    csv_file, paths_dir = get_default_csv()
except:
    UnboundLocalError

button0 = CTkButton(
    gui,
    text="Wybierz csv",
    width=160,
    height=40,
    font=("Consolas", 16),
    command=lambda v=1: ask_for_dir(v),
)
button0.pack(pady=(5, 25), anchor=CENTER)

button_csv = CTkButton(
    gui,
    text="Wybierz bugi",
    width=160,
    height=40,
    font=("Consolas", 16),
    command=lambda : open_bugs(csv_file),
)
button_csv.pack(pady=(5, 25), anchor=CENTER)

if csv_file == None:
    button_csv.configure(state='disabled')
    


button_c = CTkButton(
    gui,
    text="Wybierz pliki",
    width=160,
    height=40,
    font=("Consolas", 16),
    command=lambda v=1: open_choice_window(folder_path),
)
button_c.pack(pady=(5, 25), anchor=CENTER)
button_c.configure(state="disabled")

frame = CTkFrame(gui)
frame.configure(border_width=2, fg_color="transparent")
frame.pack()

r_check0 = CTkRadioButton(
    frame,
    text="Hotfix",
    variable=var_0,
    value=1,
    font=("Consolas", 14),
    command=get_and_display_path,
)
r_check0.pack(pady=20, padx=(15, 5), side="left", anchor=N)

r_check1 = CTkRadioButton(
    frame,
    text="Nowa wersja",
    variable=var_0,
    value=2,
    font=("Consolas", 14),
    command=get_and_display_path,
)
r_check1.pack(pady=20, padx=(5, 15), side="left", anchor=N)

label1 = CTkLabel(gui, text="", height=40, font=("Consolas", 14))
label1.pack(pady=15, anchor=CENTER)

button1 = CTkButton(
    gui,
    text="Wybierz paczkę",
    width=160,
    height=40,
    font=("Consolas", 16),
    command=lambda v=2: ask_for_dir(v),
)
button1.pack(pady=(5, 25), anchor=CENTER)

button2 = CTkButton(
    gui,
    text="Generuj listę",
    width=160,
    height=40,
    font=("Consolas", 16),
    command=lambda: make_list(folder_path, csv_file),
)
button2.pack(pady=(5, 25), anchor=CENTER)

button3 = CTkButton(
    gui,
    text="Utwórz paczkę",
    width=160,
    height=40,
    font=("Consolas", 16),
    command=lambda: copy_pack(folder_path)
)
button3.pack(pady=(5, 25), anchor=CENTER)
button3.configure(state="disabled", hover=True, border_width=1, border_color="gold")


button4 = CTkButton(gui, text="Wyślij email", width=160, height=40, font=("Consolas", 16), command=lambda: open_new_window(new_dir))

button4.pack(pady=(5, 25), anchor=CENTER)
button4.configure(state="disabled", hover=True)

gui.mainloop()

