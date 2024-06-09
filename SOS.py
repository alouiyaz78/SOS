# Création de la base de données

import sqlite3
import datetime
# Adaptateurs
def adapt_date_iso(val):
    """Adapte datetime.date au format ISO 8601."""
    return val.isoformat()

# Convertisseurs
def convert_date(val):
    """Convertit une date ISO 8601 en objet datetime.date."""
    return datetime.strptime(val.decode(), "%d/%m/%Y").date()
# Enregistrement des adaptateurs et convertisseurs
sqlite3.register_adapter(datetime.date, adapt_date_iso)
sqlite3.register_converter("date", convert_date)

conn = sqlite3.connect('database.db',detect_types=sqlite3.PARSE_DECLTYPES)
c = conn.cursor()

c.execute('''
    CREATE TABLE IF NOT EXISTS clients (
        client_id INTEGER PRIMARY KEY AUTOINCREMENT,
        Nom TEXT,
        Prenom TEXT,
        CIN TEXT,
        date_delivrance_cin TEXT,
        Date_naissance DATE,
        adresse TEXT,
        delegation TEXT,
        numero_telephone TEXT, 
        secteur_activite TEXT, 
        sous_secteur TEXT 
    )
''')

# Création de la table credits
c.execute('''
    CREATE TABLE IF NOT EXISTS credits (
        credit_id TEXT PRIMARY KEY ,
        Nom TEXT,
        Prenom TEXT,
        CIN TEXT,
        Date_Credit DATE,
        montant INTEGER,
        duree INTEGER,
        client_id INTEGER,
        FOREIGN KEY (client_id) REFERENCES clients (client_id)
    )
''')

# Creation de la table amortissemet
c.execute('''
    CREATE TABLE IF NOT EXISTS amortissement (
        reference_amortization INTEGER PRIMARY KEY AUTOINCREMENT ,
        echeance_date DATE,
        echeance_numero INTEGER,
        Montant_echeance INTEGER,
         Interet INTEGER,
         Reste_du_credit INTEGER,
         credit_id TEXT,
         date_paye DATE,
         paye TEXT,
         FOREIGN KEY (credit_id) REFERENCES credits (credit_id) ON DELETE CASCADE
    )
''')
# Création de la table de paiements
c.execute('''
    CREATE TABLE IF NOT EXISTS payements (
        reference_payement INTEGER PRIMARY KEY AUTOINCREMENT,
        paye_date DATE,
        Montant_payement INTEGER,
        Montant_reste INTEGER,
        Montant_impaye INTEGER,
        Recu TEXT,
        Mode_paiement TEXT,
        Payement_partiel INTEGER,
        credit_id TEXT,
        echeance_numero INTEGER,
        FOREIGN KEY (credit_id) REFERENCES credits (credit_id) ON DELETE CASCADE

    )
''')
c.execute('''
    CREATE TABLE IF NOT EXISTS paiements_total (
        Nom TEXT,
        Prenom TEXT,
        CIN TEXT,
        Date_Credit DATE ,
        Montant_credit INTEGER,
        Total_echeance INTEGER,
        Total_paiements INTEGER,
        Impayes INTEGER,
        credit_id TEXT,
        FOREIGN KEY (credit_id) REFERENCES credits (credit_id) ON DELETE CASCADE
    )
''')
# Supprimer toutes les lignes existantes dans la table paiements_total
c.execute('''DELETE FROM paiements_total''')

# Création de la table paiements_total avec les données mises à jour
c.execute('''
    CREATE TABLE IF NOT EXISTS paiements_total (
        Nom TEXT,
        Prenom TEXT,
        CIN TEXT,
        Date_Credit DATE,
        Montant_credit INTEGER,
        Total_echeance INTEGER,
        Total_paiements INTEGER,
        Impayes INTEGER,
        credit_id TEXT,
        FOREIGN KEY (credit_id) REFERENCES credits (credit_id) ON DELETE CASCADE
    )
''')
# Récupérer la date actuelle
aujourdhui = datetime.datetime.now().date()

# Insérer les données mises à jour dans la table paiements_total
c.execute('''
    INSERT INTO paiements_total (Nom, Prenom, CIN, Date_Credit,Montant_credit, Total_echeance, Total_paiements, Impayes, credit_id)
    SELECT DISTINCT
        cr.Nom,
        cr.Prenom,
        cr.CIN,
        cr.Date_Credit,
        cr.montant AS Montant_credit,
        am.Total_echeance,
        COALESCE(p.Total_paiements, 0) AS Total_paiements,
        COALESCE((SELECT SUM(Montant_echeance) FROM amortissement WHERE credit_id = cr.credit_id AND echeance_date < ? AND paye = ''), 0) AS Impayes,
        cr.credit_id
    FROM
        credits cr
    LEFT JOIN (
        SELECT credit_id, SUM(Montant_echeance) AS Total_echeance
        FROM amortissement
        GROUP BY credit_id
    ) am ON cr.credit_id = am.credit_id
    LEFT JOIN (
        SELECT credit_id, SUM(Montant_payement) AS Total_paiements
        FROM payements
        GROUP BY credit_id
    ) p ON cr.credit_id = p.credit_id
''', (aujourdhui,))

# Fermeture de la connexion
conn.commit()
conn.close()
from PIL import Image

Image.CUBIC = Image.BICUBIC
import tkinter as tk
from tkinter import *
from tkinter import ttk

import ttkbootstrap as tbk
from ttkbootstrap.constants import *
from ttkbootstrap import *
from tkinter import messagebox
import json
from dateutil.relativedelta import relativedelta
from datetime import datetime
import win32print
import ctypes.wintypes
import os
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib import colors, styles
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageTemplate, Preformatted, \
    PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.enums import TA_RIGHT
from bidi.algorithm import get_display
import arabic_reshaper
import io

from datetime import datetime, timedelta
from docx import Document
from docx import *
from docx.shared import Pt
from io import BytesIO
from PyPDF2 import PdfReader, PdfWriter
from arabic_reshaper import arabic_reshaper
import ctypes
import docx2pdf
import subprocess
from docx2pdf import convert
import time
import calendar
from tempfile import NamedTemporaryFile
import re
import numpy as np
import numpy_financial as npf


### Classe principale
class Application(tbk.Window):
    def __init__(self):
        super().__init__()

        self.geometry("1600x975")
        self.title("credit_sos")
        style = tbk.Style("superhero")

        self.paned_window = tbk.PanedWindow()
        self.paned_window.pack(fill=tk.BOTH, expand=True)

        menu_manager = MenuManager(self, self)  # Create an instance of MenuManager
        self.config(menu=menu_manager)  # Attach the menu to the main window

        # Create frames for IMF, Clients, and Credits sections
        self.imf_section = IMFSection(self.paned_window)
        # Create an instance of ClientsSection
        # self.clients_section = ClientsSection(self.paned_window)
        # Create an instance of PaymentSection

        self.clients_section = ClientsSection(self.paned_window, self, None)  # Pass None initially
        self.credits_section = CreditsSection(self.paned_window, self, None, None, None)
        # Create an instance of PaymentSection
        self.impression_section = ImpressionManager(self.paned_window)
        # self.impression_section=None

        # Pass clients_section as an argument to CreditsSection
        # self.credits_section = CreditsSection(self.paned_window, self, self.clients_section, self.impression_section)
        self.clients_section.credit_section = self.credits_section  # Assign credits_section after its creation
        self.payement_section = PayementSection(self.paned_window, self, self.clients_section, self.credits_section,
                                                self.impression_section)
        self.credits_section.payement_section = self.payement_section

        self.mainloop()


### Classe de creation des menus pour notre application

class MenuManager(tbk.Menu):

    def __init__(self, parent, application, **kwargs):
        super().__init__(parent, **kwargs)

        self.application = application
        self.added_sections = {'imf': False, 'clients': False, 'credits': False, 'payment': False}
        self.current_section = None

        config_menu = tbk.Menu(self, tearoff=0)
        self.add_cascade(label="Configuration", menu=config_menu)
        config_menu.add_command(label="IMF", command=self.show_imf_frame)

        payment_menu = tbk.Menu(self, tearoff=0)
        self.add_cascade(label='Recouvrements', menu=payment_menu)
        payment_menu.add_command(label='Payment', command=self.show_payment_frame)
        payment_menu.add_command(label='Recouvrement journalier', command=self.ouvrir_fenetre_rapport_journalier)
        payment_menu.add_command(label='Recouvrement Mensuel', command=self.ouvrir_fenetre_rapport_mensuel)
        payment_menu.add_command(label='Recouvrement Total',command=self.ouvrir_fenetre_rapport_gloabal)

        clients_credits_menu = tbk.Menu(self, tearoff=0)
        self.add_cascade(label="Clients/Credits", menu=clients_credits_menu)
        clients_credits_menu.add_command(label="Clients Section", command=self.show_clients_frame)
        clients_credits_menu.add_command(label="Credits Section", command=self.show_credits_frame)

    def show_imf_frame(self):
        self.destroy_current_section()
        if not self.added_sections['imf']:
            self.application.paned_window.add(self.application.imf_section)
            self.added_sections['imf'] = True
        self.application.paned_window.pack(fill=tk.BOTH, expand=True)
        self.current_section = 'imf'

    def show_clients_frame(self):
        self.destroy_current_section()
        if not self.added_sections['clients']:
            self.application.paned_window.add(self.application.clients_section)
            self.added_sections['clients'] = True
        self.application.paned_window.pack(fill=tk.BOTH, expand=True)
        self.current_section = 'clients'

    def show_credits_frame(self):
        self.destroy_current_section()
        if not self.added_sections['credits']:
            self.application.paned_window.add(self.application.credits_section)
            self.added_sections['credits'] = True
        self.application.paned_window.pack(fill=tk.BOTH, expand=True)
        self.current_section = 'credits'

    def show_payment_frame(self):
        self.destroy_current_section()
        if not self.added_sections['payment']:
            self.application.paned_window.add(self.application.payement_section)
            self.added_sections['payment'] = True
        self.application.paned_window.pack(fill=tk.BOTH, expand=True)
        self.current_section = 'payment'

    def destroy_current_section(self):
        if self.current_section and self.added_sections[self.current_section]:
            # Destroy the current section if it exists
            self.application.paned_window.forget(self.application.paned_window.panes()[0])
            self.added_sections[self.current_section] = False

    def ouvrir_fenetre_rapport_journalier(self):
        self.application.impression_section.generer_rapport_journalier()

    def ouvrir_fenetre_rapport_mensuel(self):
        self.application.impression_section.generer_rapport_mensuel()
    def ouvrir_fenetre_rapport_gloabal(self):
        self.application.impression_section.generer_rapport_global()

### CLASSE CONTENANT LES INFORMATION SUR IMF

class IMFSection(tbk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.create_widgets()
        self.load_imf_data()

    def create_widgets(self):

        frame_imf = tbk.Frame(self, bootstyle=INFO)
        frame_imf.pack(fill='both', expand=True)

        # Entry variables imf
        self.raison_sociale_imf = tbk.StringVar()
        self.adresse_imf = tbk.StringVar()
        self.autre_adresse_imf = tbk.StringVar()
        self.tel_imf = tbk.StringVar()
        self.banque_imf = tbk.StringVar()
        self.autre_banque_imf = tbk.StringVar()
        self.rib_imf = tbk.StringVar()
        self.autre_rib_imf = tbk.StringVar()
        self.cnss_imf = tbk.StringVar()
        self.rne_imf = tbk.StringVar()
        self.mail_imf = tbk.StringVar()

        # Labels
        labels_imf_arab = ['الإسم الإجتماعي', 'العنوان', ' الهاتف', 'البنك', 'رقم الحساب', 'الضمان الإجتماعي',
                           'السجل التجاري', 'العنوان الإلكتروني']
        labels_imf = [get_display(arabic_reshaper.reshape(text)) for text in labels_imf_arab]

        for i, labels_imf_text in enumerate(labels_imf):
            label = tbk.Label(frame_imf, text=labels_imf_text, style="info.Inverse.TLabel")
            label.grid(column=5, row=i, sticky=tbk.E, padx=20, pady=5)  # Alignement à droite
            label.configure(anchor='e')
            # Entry widgets
            entries_imf = [
                tbk.Entry(frame_imf, textvariable=self.raison_sociale_imf, width=40),
                tbk.Entry(frame_imf, textvariable=self.adresse_imf, width=40),
                tbk.Entry(frame_imf, textvariable=self.tel_imf, width=40),
                tbk.Entry(frame_imf, textvariable=self.banque_imf, width=40),
                tbk.Entry(frame_imf, textvariable=self.rib_imf, width=40),
                tbk.Entry(frame_imf, textvariable=self.cnss_imf, width=40),
                tbk.Entry(frame_imf, textvariable=self.rne_imf, width=40),
                tbk.Entry(frame_imf, textvariable=self.mail_imf, width=40)
            ]

        for i, entry in enumerate(entries_imf):
            entry.grid(column=4, row=i, padx=20, pady=20, sticky='w')
        # Nouveaux champs pour l'autre adresse, autre banque, et autre RIB
        tbk.Entry(frame_imf, textvariable=self.autre_adresse_imf, width=40).grid(column=2, row=1, padx=80, pady=20,
                                                                                 columnspan=2, sticky='w')
        tbk.Entry(frame_imf, textvariable=self.autre_banque_imf, width=40).grid(column=2, row=3, padx=80, pady=20,
                                                                                columnspan=2, sticky='w')
        tbk.Entry(frame_imf, textvariable=self.autre_rib_imf, width=40).grid(column=2, row=4, padx=80, pady=20,
                                                                             columnspan=2, sticky='w')
        tbk.Button(frame_imf, text="Enregistrer", style="success.TButton", width=30,
                   command=self.insert_imf).grid(column=1, row=len(entries_imf) + 1, columnspan=8, pady=25)
        # self.notebook.add(tbk.frame(self.notebook), text="IMF")

    def insert_imf(self, raison_sociale_imf='', adresse_imf='', autre_adresse_imf='', tel_imf='', banque_imf='',
                   autre_banque_imf='',
                   rib_imf='', autre_rib_imf='', cnss_imf='', rne_imf='', mail_imf=''):
        # Utiliser les valeurs actuelles si les champs sont vides
        raison_sociale_imf = self.raison_sociale_imf.get()
        adresse_imf = self.adresse_imf.get()
        autre_adresse_imf = self.autre_adresse_imf.get()
        tel_imf = self.tel_imf.get()
        banque_imf = self.banque_imf.get()
        autre_banque_imf = self.autre_banque_imf.get()
        rib_imf = self.rib_imf.get()
        autre_rib_imf = self.autre_rib_imf.get()
        cnss_imf = self.cnss_imf.get()
        rne_imf = self.rne_imf.get()
        mail_imf = self.mail_imf.get()

        imf_data = {
            "Raison Sociale": raison_sociale_imf,
            "Adresse": {"Compte 1": adresse_imf, "Compte 2": autre_adresse_imf},
            "Tel": tel_imf,
            "Banque": {"Compte 1": banque_imf, "Compte 2": autre_banque_imf},
            "RIB": {"Compte 1": rib_imf, "Compte 2": autre_rib_imf},
            "CNSS": cnss_imf,
            "RNE": rne_imf,
            "MAIL": mail_imf
        }

        json_file_path = "donnees_imf.json"
        with open(json_file_path, "w") as json_file:
            json.dump(imf_data, json_file, indent=4)

        # Boîte de message de confirmation
        messagebox.showinfo("Succès", f"Les données IMF ont été enregistrées dans {json_file_path}")

        # Mettre à jour les valeurs actuelles des champs IMF
        self.raison_sociale_imf_value = raison_sociale_imf
        self.adresse_imf_value = adresse_imf
        self.autre_adresse_imf_value = autre_adresse_imf
        self.tel_imf_value = tel_imf
        self.banque_imf_value = banque_imf
        self.autre_banque_imf_value = autre_banque_imf
        self.rib_imf_value = rib_imf
        self.autre_rib_imf_value = autre_rib_imf
        self.cnss_imf_value = cnss_imf
        self.rne_imf_value = rne_imf
        self.mail_imf_value = mail_imf

    def load_imf_data(self):
        json_file_path = "donnees_imf.json"
        try:
            with open(json_file_path, "r") as json_file:
                imf_data = json.load(json_file)

            # Remplir les champs avec les données du fichier JSON
            self.raison_sociale_imf.set(imf_data.get("Raison Sociale", ""))
            self.adresse_imf.set(imf_data.get("Adresse", {}).get("Compte 1", ""))
            self.autre_adresse_imf.set(imf_data.get("Adresse", {}).get("Compte 2", ""))
            self.tel_imf.set(imf_data.get("Tel", ""))
            self.banque_imf.set(imf_data.get("Banque", {}).get("Compte 1", ""))
            self.autre_banque_imf.set(imf_data.get("Banque", {}).get("Compte 2", ""))
            self.rib_imf.set(imf_data.get("RIB", {}).get("Compte 1", ""))
            self.autre_rib_imf.set(imf_data.get("RIB", {}).get("Compte 2", ""))
            self.cnss_imf.set(imf_data.get("CNSS", ""))
            self.rne_imf.set(imf_data.get("RNE", ""))
            self.mail_imf.set(imf_data.get("MAIL", ""))
        except FileNotFoundError:
            # Gérer le cas où le fichier JSON n'existe pas encore
            print(f"Le fichier JSON {json_file_path} n'existe pas encore.")
        except json.JSONDecodeError as e:
            # Gérer le cas où il y a une erreur de décodage JSON
            print(f"Erreur de décodage JSON : {e}")

            # Section clients : information sur les clients


class ClientsSection(tbk.Frame):
    def __init__(self, parent, application, credit_section, **kwargs):
        super().__init__(parent, **kwargs)
        self.create_widgets()
        self.application = application
        self.credit_section = credit_section

    def create_widgets(self):
        frame_clients = tbk.Frame(self, bootstyle=INFO)
        frame_clients.pack(fill='both', expand=True)

        # Utilisez StringVar() pour les variables des champs de saisie
        Nom_var = tk.StringVar()
        Prenom_var = tk.StringVar()
        cin_var = tk.StringVar()
        date_delivrance_cin_var = tbk.DateEntry(frame_clients)

        date_de_naissance_var = tbk.DateEntry(frame_clients)

        adresse_var = tk.StringVar()
        delegation_var = tk.StringVar()
        numero_telephone_var = tk.StringVar()
        secteur_activite_var = tk.StringVar()
        sous_secteur_var = tk.StringVar()

        labels = ['Nom', 'Prenom', 'CIN', 'Date de délivrance CIN', 'Date de naissance',
                  'Adresse', 'Délégation', 'Numéro de téléphone', 'Secteur d\'activité', 'Sous-secteur']

        for i, label_text in enumerate(labels):
            tbk.Label(frame_clients, text=label_text, style="info.Inverse.TLabel").grid(column=0, row=i,
                                                                                        sticky=tbk.W,
                                                                                        padx=20, pady=5)

        entries = [
            tbk.Entry(frame_clients, textvariable=Nom_var, width=40),
            tbk.Entry(frame_clients, textvariable=Prenom_var, width=40),
            tbk.Entry(frame_clients, textvariable=cin_var, width=30),
            date_delivrance_cin_var,
            date_de_naissance_var,
            tbk.Entry(frame_clients, textvariable=adresse_var, width=40),
            tbk.Entry(frame_clients, textvariable=delegation_var, width=40),
            tbk.Entry(frame_clients, textvariable=numero_telephone_var, width=30),
            tbk.Entry(frame_clients, textvariable=secteur_activite_var, width=40),
            tbk.Entry(frame_clients, textvariable=sous_secteur_var, width=40)
        ]

        for i, entry in enumerate(entries):
            entry.grid(column=1, row=i, columnspan=4, padx=20, pady=10)

        tbk.Button(frame_clients, text="Ajouter Client", style="success.TButton", width=30,
                   command=lambda: self.insert_client(
                       Nom_var.get(),
                       Prenom_var.get(),
                       cin_var.get(),
                       date_delivrance_cin_var.entry.get(),
                       date_de_naissance_var.entry.get(),
                       adresse_var.get(),
                       delegation_var.get(),
                       numero_telephone_var.get(),
                       secteur_activite_var.get(),
                       sous_secteur_var.get()
                   )).grid(column=1, row=len(entries) + 1, columnspan=8, pady=25)

    def insert_client(self, Nom, Prenom, cin, date_delivrance_cin, Date_naissance, adresse, delegation,
                      numero_telephone, secteur_activite, sous_secteur):

        conn = sqlite3.connect('database.db')
        c = conn.cursor()

        c.execute('''
                  INSERT INTO clients (Nom, Prenom, cin, date_delivrance_cin, Date_naissance,
                                       adresse, delegation, numero_telephone, secteur_activite, sous_secteur)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                  ''', (Nom, Prenom, cin, date_delivrance_cin, Date_naissance, adresse, delegation,
                        numero_telephone, secteur_activite, sous_secteur))

        conn.commit()  # Actualiser la liste des clients dans l'interface utilisateur
        conn.close()
        clients = self.credit_section.get_clients()
        self.credit_section.recherche_combo['values'] = clients
        print(clients)  # Récupérer la liste des clients

        messagebox.showinfo("Succès", "Client ajouté avec succès!")

    # CLASSE CREDITS : CREATION DES CREDITS POUR LES CLIENTS /TABLEAUX D'AMORTISSEMENT


class CreditsSection(tbk.Frame):
    credit_counter = 0
    amortization_data = None

    def __init__(self, parent, application, clients_section, impression_section, payement_section, **kwargs):
        super().__init__(parent, **kwargs)
        self.application = application
        self.clients_section = clients_section
        self.impression_section = impression_section
        self.payement_section = payement_section

        # self.amortization_table = None
        self.create_widgets()
        self.amortization_data = None
        self.amortization_data_temporaire = None

    def create_widgets(self, amortization_data=None):
        frame_credits = tbk.Frame(self, bootstyle=INFO)
        frame_credits.pack(fill='both', expand=True)

        self.Nom_var_credit = tk.StringVar()
        self.Prenom_var_credit = tk.StringVar()
        self.cin_var_credit = tk.StringVar()
        self.credit_id_var_credit = tk.StringVar()
        self.taux_interet_var = tk.DoubleVar()
        self.taux_interet_label = tk.Label(frame_credits, text="", anchor="w")
        self.Date_Credit = tbk.DateEntry(frame_credits)
        montant_var = tk.IntVar()
        duree_var = tk.IntVar()
        grace_var = tk.IntVar()
        date_Credit_var = tk.StringVar()
        # self.duree_grace=tk.IntVar()
        self.grace_var = tk.IntVar()
        self.duree_values = tk.IntVar()
        self.mensualite = tk.IntVar()
        self.clients_var = StringVar()

        duree_values = list(range(1, 37))
        duree_dropdown = tbk.Combobox(frame_credits, textvariable=duree_var, values=duree_values, state="readonly")
        duree_dropdown.set(duree_values[0])
        duree_grace_val = list(range(0, 13))
        duree_grace = tbk.Combobox(frame_credits, textvariable=self.grace_var, values=duree_grace_val, state="readonly")
        duree_grace.set(duree_grace_val[0])
        label_recherche = tbk.Label(frame_credits, text='Recherche Client')
        label_recherche.grid(column=0, row=0, sticky=tbk.W, padx=10, pady=5)
        self.recherche_combo = tbk.Combobox(frame_credits, textvariable=self.clients_var, state="readonly",
                                            bootstyle="success",
                                            width=30)
        self.recherche_combo.grid(column=1, row=0, sticky=tbk.W, padx=10, pady=5)

        self.recherche_combo['values'] = self.get_clients()

        self.recherche_combo.bind("<<ComboboxSelected>>",
                                  lambda event, dropdown=self.recherche_combo: self.on_client_selected(dropdown.get()))
        labels_credits = ['Nom', 'Prenom', 'CIN', 'Date_Credit', 'Montant (DT)', 'Durée (en mois)', 'Duré de grace',
                          'Crédit ID', 'Taux d"interet', '', 'Mensualités']

        for i, label_text in enumerate(labels_credits):
            tbk.Label(frame_credits, text=label_text).grid(column=0, row=i + 1, sticky=tbk.W, padx=10, pady=5)

        entries_credits = [
            tbk.Entry(frame_credits, textvariable=self.Nom_var_credit, width=40),
            tbk.Entry(frame_credits, textvariable=self.Prenom_var_credit, width=40),
            tbk.Entry(frame_credits, textvariable=self.cin_var_credit, width=30, justify='center'),
            self.Date_Credit,
            tbk.Entry(frame_credits, textvariable=montant_var, width=20),
            duree_dropdown,
            duree_grace,
            tbk.Entry(frame_credits, textvariable=self.credit_id_var_credit, state='readonly'),

            tbk.Scale(frame_credits, variable=self.taux_interet_var, bootstyle="warning", length=200, from_=0, to=30,
                      orient=tk.HORIZONTAL,
                      command=self.update_taux_interet_label),
            self.taux_interet_label,
            tbk.Entry(frame_credits, textvariable=self.mensualite, width=30, justify='center')
        ]

        for i, entry in enumerate(entries_credits):
            entry.grid(column=1, row=i + 1, padx=20, pady=10)
        duree_values = list(range(1, 37))

        tbk.Button(frame_credits, text="Tableaux d'amortissement", style="success.TButton", width=30,
                   command=lambda: self.afficher_credit(
                       self.Nom_var_credit.get(),
                       self.Prenom_var_credit.get(),
                       self.cin_var_credit.get(),
                       date_Credit_var.get(),
                       montant_var.get(),
                       duree_dropdown.get(),
                       self.credit_id_var_credit.get()
                   )).grid(column=1, row=len(entries_credits) + 1, columnspan=8, pady=25, sticky='w')
        # Create columns for the Treeview
        style = tbk.Style()
        style.configure("mystyle.Treeview", highlightthickness=1, bd=0,
                        font=('Calibri', 11))  # Modify the font of the body
        style.configure("mystyle.Treeview.Heading", font=('Calibri', 12, 'bold'),
                        background="#808080")  # Modify the font of the headings

        self.amortization_table = tbk.Treeview(frame_credits, style="mystyle.Treeview", columns=(
            'Date', 'Numéro d\'échéance', 'Mensualité', 'Intérêt', 'Reste du crédit'))

        # Define column headings
        self.amortization_table.heading('#0', text='', anchor='w')  # Placeholder for index column
        self.amortization_table.column('#0', width=1, anchor='w')
        self.amortization_table.heading('Date', text='Date', anchor='w')
        self.amortization_table.heading('Numéro d\'échéance', text='Numéro d\'échéance', anchor='w')
        self.amortization_table.heading('Mensualité', text='Mensualité', anchor='w')
        self.amortization_table.heading('Intérêt', text='Intérêt', anchor='w')
        self.amortization_table.heading('Reste du crédit', text='Reste du crédit', anchor='w')
        self.amortization_table.grid(row=0, column=2, rowspan=len(entries_credits) + 2, columnspan=6, sticky='nsew')

        self.amortization_table.insert(parent='', index="end")

        btn_valider = tbk.Button(frame_credits, text="Valider", style="success.TButton",
                                 command=lambda: self.valider_credit(self.Nom_var_credit.get(),
                                                                     self.Prenom_var_credit.get(),
                                                                     self.cin_var_credit.get(),
                                                                     self.Date_Credit.entry.get(), montant_var.get(),
                                                                     duree_dropdown.get(), amortization_data))

        btn_imprimer = tbk.Button(frame_credits, text="Imprimer Table", style="success.TButton",
                                  command=lambda : self.application.impression_section.generate_and_print_table(
                                      self.Nom_var_credit.get(), self.Prenom_var_credit.get(),
                                      self.cin_var_credit.get()))
        btn_contrat = tbk.Button(frame_credits, text="Imprimer contrat", style="Success", width=30,
                                 command=lambda: self.application.impression_section.generate_and_print_contrat(
                                     self.Nom_var_credit.get(), self.Prenom_var_credit.get(),
                                     self.cin_var_credit.get()))
        btn_traite = tbk.Button(frame_credits, text="Imprimer traites", style="danger", width=40,
                                command=lambda: self.application.impression_section.generate_and_print_traites(
                                    self.Nom_var_credit.get(), self.Prenom_var_credit.get(), self.cin_var_credit.get()))

        # Ajouter les boutons en dessous du Treeview
        btn_valider.grid(row=len(entries_credits) + 2, column=0, padx=20, pady=30)
        btn_imprimer.grid(row=len(entries_credits) + 2, column=1, padx=20, pady=30)
        btn_contrat.grid(row=len(entries_credits) + 2, column=3, padx=20, pady=30)
        btn_traite.grid(row=len(entries_credits) + 2, column=4, padx=20, pady=30)
        # Ajoutez l'instance d'AmortizationTable et utilisez grid

        # Pour permettre le redimensionnement des colonnes/ lignes
        frame_credits.columnconfigure(2, weight=1)
        frame_credits.rowconfigure(len(entries_credits) + 2, weight=1)

    def get_clients(self):
        conn = sqlite3.connect('database.db')
        c = conn.cursor()

        c.execute('SELECT Nom, Prenom FROM clients')
        clients = [" ".join(row) for row in c.fetchall()]

        conn.close()
        return clients

    def on_client_selected(self, selected_client):
        if selected_client:
            conn = sqlite3.connect('database.db')
            c = conn.cursor()

            c.execute('''
                SELECT cin, Nom, Prenom FROM clients 
                WHERE (Nom || ' ' || Prenom) = ? 
                ORDER BY client_id DESC 
                LIMIT 1
            ''', (selected_client,))

            result = c.fetchone()

            if result:
                cin, Nom, Prenom = result
                print("Client trouvé :", Nom, Prenom, cin)

                self.Nom_var_credit.set(Nom)
                self.Prenom_var_credit.set(Prenom)
                self.cin_var_credit.set(cin)

                self.credit_id_var_credit.set(self.generate_credit_id())
            else:
                print("Aucun client trouvé avec le Nom et le Prenom spécifiés.")

            conn.close()

    def generate_credit_id(self):
        CreditsSection.credit_counter += 1
        return f"CR{CreditsSection.credit_counter}"

    def update_taux_interet_label(self, value):
        rounded_value = round(float(value))  # Arrondir à la valeur entière la plus proche
        self.taux_interet_var.set(rounded_value)
        self.taux_interet_label.config(text=f"Taux d'intérêt : {rounded_value} %")

    def afficher_credit(self, Nom, Prenom, CIN, Date_Credit, montant, duree, credit_id):
        # Génère les données d'amortissement
        date_credit_str = self.Date_Credit.entry.get()
        if date_credit_str:
            # Convertir la chaîne de date en objet datetime
            date_format = '%d/%m/%Y'
            date_credit = datetime.strptime(date_credit_str, date_format)

            amortization_data = self.calculate_amortization_data(
                int(montant), int(duree), self.taux_interet_var.get(),
                date_credit, self.grace_var.get()
            )
            # Stocker temporairement les données d'amortissement dans la variable de classe
            self.amortization_data_temporaire = amortization_data

            # Mettre à jour et afficher le tableau d'amortissement
            self.update_amortization_table(amortization_data)
        else:
            messagebox.showerror("Erreur", "Il y'a des champs nom remplis.")

    def calculate_amortization_data(self, montant, duree, taux_interet, date_credit, grace):
        amortization_data = []
        total_echeances = 0
        montant_echeance = 0
        taux_interet_decimal = round(taux_interet / 100, 3)
        # Calcul de la durée réelle du crédit en tenant compte de la durée de grâce
        duree_reelle = duree + grace

        # Vérification si la durée réelle du crédit dépasse 3 ans (36 mois)
        if duree_reelle > 36:
            # Affichage d'un message d'alerte avec Tkinter
            root = tk.Tk()
            root.withdraw()  # Pour cacher la fenêtre principale
            messagebox.showwarning("Durée du crédit trop longue",
                                   "La durée totale du crédit (y compris la période de grâce) ne doit pas dépasser 3 ans (36 mois).")
            root.destroy()  # Pour détruire la fenêtre Tkinter après l'affichage du message
            return None

        # Ajouter une condition pour gérer le cas où le taux d'intérêt est nul
        if taux_interet == 0:
            # Calculer la mensualité en utilisant la formule du prêt simple
            mensualite = montant / duree

            # La date de la première échéance est la date de crédit + durée de grâce
            date_echeance = date_credit + relativedelta(months=grace + 1)

            reste_credit = montant  # Initialiser le reste du crédit

            for numero_echeance in range(1, duree + 1):
                # Ajouter les données à la liste
                if numero_echeance == duree:
                    # Dernière échéance, ajouter le reste à la mensualité
                    montant_echeance = reste_credit
                else:
                    # Arrondir à l'entier multiple de 10 le plus proche
                    montant_echeance = round(mensualite / 10) * 10

                amortization_data.append((
                    date_echeance.strftime('%Y-%m-%d'),
                    numero_echeance,
                    f"{montant_echeance:.0f} DT",
                    "0 DT",  # Aucun intérêt puisque le taux est nul
                    f"{(reste_credit):.0f} DT"  # Reste du crédit
                ))

                # Soustraire le montant de l'échéance du reste du crédit
                reste_credit -= montant_echeance
                total_echeances += montant_echeance

                # Augmenter la date pour la prochaine échéance
                date_echeance += relativedelta(months=1)



        else:

            # Initialiser les valeurs à l'extérieur de la boucle
            reste_credit = montant
            montant_echeance_constante = round((montant / duree) / 10) * 10  # Calcul de l'échéance constante
            # La date de la première échéance est la date de crédit + durée de grâce
            date_echeance = date_credit + relativedelta(months=grace + 1)

            # Initialiser le facteur multiplicatif des intérêts
            facteur_interet = 1

            for numero_echeance in range(1, duree + 1):
                montant_echeance = 0  # Initialiser le montant de l'échéance à zéro

                # Si la période de grâce est terminée, ajouter les échéances normales
                if numero_echeance > grace:
                    montant_echeance = montant_echeance_constante
                    date_echeance += relativedelta(months=1)  # Mettre à jour la date d'échéance

                # Mettre à jour le facteur multiplicatif si la période de grâce n'est pas terminée
                if numero_echeance <= grace:
                    facteur_interet *= 2  # Double le facteur d'intérêt

                # Recalculer les intérêts en utilisant le nouveau facteur
                interet = round((reste_credit * taux_interet_decimal * facteur_interet) / 12, 3)

                # Ajouter l'échéance à la liste
                amortization_data.append((
                    date_echeance.strftime('%Y-%m-%d'),
                    numero_echeance,
                    f"{montant_echeance:.0f} DT",
                    f"{interet:.0f} DT",
                    f"{reste_credit:.0f} DT"
                ))

                # Mettre à jour le reste du crédit pour les échéances normales
                if numero_echeance > grace:
                    reste_credit = reste_credit - montant_echeance + interet

                total_echeances += montant_echeance

            # Calcul de l'échéance pour la dernière échéance
            montant_derniere_echeance = montant - (montant_echeance_constante * (duree - grace))

            # Mettre à jour la dernière échéance
            amortization_data[-1] = (
                amortization_data[-1][0],  # Conserver la date
                amortization_data[-1][1],  # Conserver le numéro d'échéance
                f"{montant_derniere_echeance:.0f} DT",  # Utiliser le montant calculé pour la dernière échéance
                f"{interet:.0f} DT",  # Utiliser le dernier calcul d'intérêt
                amortization_data[-1][4]  # Conserver le reste du crédit
            )
        print(f"La somme totale des échéances est de : {total_echeances:.0f} DT")

        return amortization_data

    def clear_amortization_table(self):
        for item in self.amortization_table.get_children():
            self.amortization_table.delete(item)

    def update_amortization_table(self, amortization_data):
        # Effacez le contenu actuel du tableau
        self.clear_amortization_table()

        # Ajoutez les nouvelles données du tableau
        for row in amortization_data:
            self.amortization_table.insert('', 'end', values=row)

    def valider_credit(self, Nom, Prenom, CIN, Date_Credit, montant, duree, amortization_data):
        conn = sqlite3.connect('database.db')
        c = conn.cursor()

        try:
            # Commencer une transaction
            conn.execute('BEGIN')

            # Recherche du client_id correspondant aux valeurs de Nom, Prenom, et CIN dans la table clients
            c.execute('''
                SELECT client_id FROM clients 
                WHERE Nom = ? AND Prenom = ? AND CIN = ?
            ''', (Nom, Prenom, CIN))

            result_1 = c.fetchone()

            if result_1:
                client_id = result_1[0]

                # Générer une référence unique pour le crédit
                reference_credit = self.generer_reference_unique("CR")

                # Insérer les données dans la table credits
                c.execute('''
                    INSERT INTO credits (credit_id, Nom, Prenom, CIN, Date_Credit, montant, duree, client_id)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (reference_credit, Nom, Prenom, CIN, Date_Credit, montant, duree, client_id))

                # Récupérer l'ID du crédit nouvellement inséré
                credit_id = c.lastrowid

                # Utiliser les données d'amortissement temporairement stockées
                if self.amortization_data_temporaire:
                    amortization_data = self.amortization_data_temporaire

                    for data in amortization_data:
                        # Assurez-vous que 'data' a toutes les colonnes nécessaires dans l'ordre correct
                        echeance_date, echeance_numero, Montant_echeance, Interet, Reste_du_credit = data

                        # Laisser les nouvelles colonnes vides pour le moment
                        paye = ""  # Nouvelle colonne "paye"
                        date_paye = ""  # Nouvelle colonne "date_paye"

                        # Insérer les données dans la table d'amortissement en utilisant le credit_id récupéré
                        c.execute('''
                            INSERT INTO amortissement (echeance_date, echeance_numero, Montant_echeance, Interet, Reste_du_credit, paye, date_paye, credit_id)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            echeance_date, echeance_numero, Montant_echeance, Interet, Reste_du_credit, paye, date_paye,
                            reference_credit))

            # Valider la transaction
            conn.execute('COMMIT')

            messagebox.showinfo("Succès", "Crédit ajouté avec succès!")
            # mis à jour et synchronisation de la liste des credits
            self.payement_section.update_client_list()

            # Réinitialiser la variable de classe après utilisation
            self.amortization_data_temporaire = None

        except Exception as e:
            # En cas d'erreur, annuler la transaction
            conn.execute('ROLLBACK')
            messagebox.showerror("Erreur", f"Une erreur s'est produite : {str(e)}")

        finally:
            # Fermer la connexion
            conn.close()

    def generer_reference_unique(self, prefixe):
        conn = sqlite3.connect('database.db')
        c = conn.cursor()

        # Récupérer le dernier numéro utilisé pour le préfixe donné
        c.execute('SELECT MAX(SUBSTR(credit_id, LENGTH(?) + 1)) FROM credits WHERE credit_id LIKE ?',
                  (prefixe, f"{prefixe}%"))
        dernier_numero = c.fetchone()[0]
        dernier_numero = int(dernier_numero) if dernier_numero else 0

        # Si le dernier numéro n'est pas dans la base de données, commencez à 1
        nouveau_numero = dernier_numero + 1

        conn.close()

        # Générer la référence complète
        reference_complete = f"{prefixe}{nouveau_numero}"

        return reference_complete


### CLASSE QUI GERE LES PAYEMENTS DE CREDIT VALIDATION ET ENREGISTREMENT DES PAYEMENTS /CALCUL DES IMPAYER DU RESTE DES ENCOURS

class PayementSection(tbk.Frame):

    def __init__(self, parent, application, clients_section, credits_section, impression_section, **kwargs):
        super().__init__(parent, **kwargs)
        self.application = application
        self.clients_section = clients_section
        self.credits_section = credits_section
        self.impression_section = impression_section
        self.create_widgets()
        self.client_names = []
        self.numero_recu_initial = 1

    def create_widgets(self):
        self.taux_recouvrement = 100
        # Configuration de la grille principale
        self.rowconfigure(0, weight=1)  # Première rangée
        self.rowconfigure(1, weight=1)  # Deuxième rangée
        self.rowconfigure(2, weight=1)  # Troisième rangée
        self.rowconfigure(3, weight=1)  # Quatrième rangée
        self.rowconfigure(4, weight=1)  # Cinquième rangée
        self.rowconfigure(7, weight=1)  # Cinquième rangée
        self.rowconfigure(8, weight=1)  # Cinquième rangée
        self.columnconfigure(0, weight=1)  # Première colonne
        self.columnconfigure(1, weight=1)  # Deuxième colonne
        self.columnconfigure(2, weight=1)  # Troisième colonne
        self.columnconfigure(3, weight=1)  # Première colonne
        self.columnconfigure(4, weight=1)  # Deuxième colonn

        # Cadre principal pour les widgets de recherche
        frame_recherche = tbk.LabelFrame(self, style='Info', text='Recherche clients')
        frame_recherche.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # Labels et widgets de recherche
        recherche_nom = tbk.Label(frame_recherche, text='Rechercher dans la liste', font=('Halvetica', 12),
                                  style="secondary.Inverse.TLabel")
        recherche_nom.grid(row=0, column=0, padx=10, pady=20, sticky="nsew")

        recherche_cin = tbk.Label(frame_recherche, text='Rechercher par CIN', font=('Halvetica', 12),
                                  style='secondary.Inverse.TLabel')
        recherche_cin.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")

        self.entry_cin = tbk.Entry(frame_recherche, font=('Halvetica', 12), width=20, style='secondary')
        self.entry_cin.grid(row=2, column=1, padx=10, pady=10, sticky="nsew")
        self.entry_cin.bind('<KeyRelease>', self.filter_clients)

        self.entry = tbk.Entry(frame_recherche, font=('Halvetica', 12), width=20, style='secondary')
        self.entry.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        self.entry.bind('<KeyRelease>', self.filter_clients)

        self.credit_dropdown = tbk.Combobox(frame_recherche, width=20)
        self.credit_dropdown.grid(row=1, column=1, padx=10, pady=10, sticky="nsew")
        self.credit_dropdown.bind('<<ComboboxSelected>>', self.update_amortization_table)

        self.update_client_list()

        # Créer une Frame pour contenir la table d'amortissement
        frame_table = tbk.Frame(self, style='Info')
        frame_table.grid(row=0, column=1, padx=10, columnspan=6, rowspan=4, sticky="nsew")

        # Créer une Frame pour la table d'amortissement à l'intérieur de la Frame précédente
        frame_tab = tbk.LabelFrame(frame_table, style='Info', text='Tableau_credit')
        frame_tab.pack(fill="both", expand=True)  # Utilise pack pour remplir la Frame

        frame_tab.columnconfigure(0, weight=1)
        frame_tab.rowconfigure(0, weight=1)
        style = tbk.Style()
        style.configure("mystyle.Treeview", highlightthickness=1, bd=0,
                        font=('Calibri', 11))  # Modify the font of the body
        style.configure("mystyle.Treeview.Heading", font=('Calibri', 12, 'bold'),
                        background="#808080")  # Modify the font of the headings

        # style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])  # Remove the borders

        self.amortization_table = tbk.Treeview(frame_tab, columns=(
            'Date', 'N°', 'Mensualité', 'Reste', 'Intérêt', 'Payer', 'Date paye'), style="mystyle.Treeview")

        # Définir la largeur des colonnes
        self.amortization_table.heading('#0', text='', anchor='w')
        self.amortization_table.column('#0', width=1, anchor='w')

        # Réduire l'espacement entre les colonnes du tableau
        self.amortization_table['padding'] = (5, 0)  # (horizontal, vertical)

        self.amortization_table.heading('Date', text='Date', anchor='w')
        self.amortization_table.column('Date', width=25, anchor='w')  # Ajustez la largeur selon vos besoins

        self.amortization_table.heading('N°', text='N°', anchor='w')
        self.amortization_table.column('N°', width=10, anchor='w')  # Ajustez la largeur selon vos besoins

        self.amortization_table.heading('Mensualité', text='Mensualité', anchor='w')
        self.amortization_table.column('Mensualité', width=25)  # Ajustez la largeur selon vos besoins

        self.amortization_table.heading('Reste', text='Reste', anchor='w')
        self.amortization_table.column('Reste', width=25)  # Ajustez la largeur selon vos besoins
        self.amortization_table.heading('Intérêt', text='Interet', anchor='w')
        self.amortization_table.column('Intérêt', width=25)  # Ajustez la largeur selon vos besoins

        self.amortization_table.heading('Payer', text='Payer', anchor='w')
        self.amortization_table.column('Payer', width=10)
        self.amortization_table.heading('Date paye', text='Date paye', anchor='w')
        self.amortization_table.column('Date paye', width=25)
        self.amortization_table.grid(row=0, column=0, sticky="nswe")

        frame_credit = tbk.LabelFrame(self, style='Info', text='PAYEMENT')
        frame_credit.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

        # client labels
        self.info_client_Label = Label(frame_credit, text='Client', font=('Halvetica', 12),
                                       style="secondary.Inverse.TLabel")
        self.info_client_Label.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        self.entry_np = tbk.Entry(frame_credit, font=('Halvetica', 12), foreground='white', width=20, style='warning')
        self.entry_np.grid(row=2, column=1, padx=10, pady=20, sticky="nsew")
        self.entry_cr = tbk.Entry(frame_credit, font=('Halvetica', 12), foreground='white', width=20, style='warning')
        self.entry_cr.grid(row=2, column=2, padx=10, pady=10, sticky="nsew")
        # credit Label
        self.info_credit_Label = Label(frame_credit, text='Credit/Reste', font=('Halvetica', 12),
                                       style="secondary.Inverse.TLabel")
        self.info_credit_Label.grid(row=3, column=0, padx=10, pady=10, sticky="nsew")
        self.entry_cred = tbk.Entry(frame_credit, font=('Halvetica', 12), foreground='white', width=20, style='warning')
        self.entry_cred.grid(row=3, column=1, padx=10, pady=10, sticky="nsew")
        self.entry_reste = tbk.Entry(frame_credit, font=('Halvetica', 12), foreground='white', width=20,
                                     style='warning')
        self.entry_reste.grid(row=3, column=2, padx=10, pady=10, sticky="nsew")
        # self.credit_dropdown.bind("<<ComboboxSelected>>", self.update_entry_fields)
        # Payement Label
        self.paye_Label = Label(frame_credit, text='Payement', font=('Halvetica', 12),
                                style="secondary.Inverse.TLabel")
        self.paye_Label.grid(row=4, column=0, padx=10, pady=10, sticky="nsew")
        self.entry_montant = tbk.Entry(frame_credit, font=('Halvetica', 12), width=20, style='secondary')
        self.entry_montant.grid(row=4, column=1, padx=10, pady=10, sticky="nsew")
        Liste = ['Espece', 'Chéque', 'Virement', 'En ligne']
        self.modalite = tbk.Combobox(frame_credit, width=20, values=Liste)
        self.modalite.grid(row=4, column=2)
        self.modalite.current(0)
        self.date_Label = Label(frame_credit, text='Date Payment', font=('Halvetica', 12),
                                style="secondary.Inverse.TLabel")
        self.date_Label.grid(row=5, column=0)
        self.date_pay = tbk.DateEntry(frame_credit, style='warning')
        self.date_pay.grid(row=5, column=1)
        self.recu_label = Label(frame_credit, text='N° Recu', font=('Halvetica', 12),
                                style="secondary.Inverse.TLabel")
        self.recu_label.grid(row=6, column=0, padx=10, pady=10, sticky="nsew")
        self.entry_recu = tbk.Entry(frame_credit, font=('Halvetica', 12), width=10, foreground='white', style='warning')
        self.entry_recu.grid(row=6, column=1, padx=10, pady=10, sticky="nsew")

        # self.entry_recu_auto = tbk.Entry(frame_credit, font=('Halvetica', 12), width=20, style='secondary')
        # self.entry_recu_auto.grid(row=6, column=2, padx=10, pady=20, sticky="nsew")
        self.valider = tbk.Button(frame_credit, width=20, style='success', text='Valider',
                                  command=self.effectuer_paiement)
        self.valider.grid(row=7, column=1, columnspan=2, padx=10, pady=10, sticky="nsew")
        self.grid_rowconfigure(3, weight=1)

        label_stat_client = tbk.LabelFrame(self, style='Info', text='STATISTIQUE')
        label_stat_client.grid(row=2, column=0, padx=10, pady=(0, 5), sticky="nsew")
        self.taux_recouv = Meter(label_stat_client, bootstyle="success", subtext='Taux Rec', textright="%",
                                 subtextstyle="light", metersize=150,
                                 stripethickness=10, textfont=('Halvetica', 10, 'bold'),
                                 amountused=self.taux_recouvrement, meterthickness=10, padding=50)
        self.taux_recouv.grid(row=0, column=1, padx=5, pady=(0, 5))
        self.titre = tbk.LabelFrame(label_stat_client, style='default', text="IMPAYE")
        self.titre.grid(row=0, column=2, padx=15)
        self.impaye_entry = tbk.Entry(self.titre, font=('Helvetica', 10), width=15, style="danger", foreground='white')
        self.impaye_entry.grid(row=2, column=1, rowspan=2, sticky="ns")
        self.titre_2 = tbk.LabelFrame(label_stat_client, style='default', text="Date ENP")
        self.titre_2.grid(row=0, column=3)
        self.date_np_entry = tbk.Entry(self.titre_2, font=('Helvetica', 10), width=15, style="danger",
                                       foreground='white')
        self.date_np_entry.grid(row=2, column=1, rowspan=2, sticky="ns")

    def update_client_list(self):
        conn = sqlite3.connect('database.db')
        c = conn.cursor()

        # Récupérer tous les clients de la base de données
        c.execute("SELECT Prenom, Nom ,credit_id FROM credits")
        results = c.fetchall()
        self.client_names = [f"{row[0]} {row[1]} - {row[2]}" for row in results]

        # Mettre à jour les options du Combobox avec tous les clients
        self.credit_dropdown['values'] = self.client_names

        c.close()

    def filter_clients(self, event=None):
        search_text_nom = self.entry.get().strip().lower()
        search_text_cin = str(self.entry_cin.get()).strip().lower()

        conn = sqlite3.connect('database.db')
        c = conn.cursor()

        if not search_text_nom and not search_text_cin:  # Si les champs de saisie sont vides, réinitialiser la liste des clients
            self.update_client_list()
        else:
            # Si le numéro de CIN est saisi, rechercher le client correspondant
            if search_text_cin:
                c.execute("SELECT Prenom, Nom, credit_id FROM credits WHERE CIN LIKE ?",
                          ('%' + search_text_cin + '%',))
                results = c.fetchall()
            else:
                # Filtrer les clients en fonction du texte saisi dans le nom
                c.execute("SELECT Prenom, Nom, credit_id FROM credits WHERE Prenom LIKE ? OR Nom LIKE ?",
                          ('%' + search_text_nom + '%', '%' + search_text_nom + '%'))
                results = c.fetchall()

            filtered_clients = [f"{row[0]} {row[1]} - {row[2]}" for row in results]  # Liste de noms complets

            # Mettre à jour les options du Combobox avec les clients filtrés
            self.credit_dropdown['values'] = filtered_clients
            self.credit_dropdown.current(0)

        c.close()

    def update_amortization_table(self, event=None):
        selected_credit = self.credit_dropdown.get()
        if not selected_credit:
            return

        # Extraire le credit_id du crédit sélectionné
        selected_credit_id = selected_credit.split(' - ')[-1]

        # Maintenant, vous pouvez utiliser selected_credit_id pour récupérer les informations
        # d'amortissement de la table correspondante dans la base de données
        conn = sqlite3.connect('database.db')
        c = conn.cursor()

        c.execute(
            "SELECT echeance_date,echeance_numero,Montant_echeance,Reste_du_credit,Interet,paye,date_paye FROM amortissement WHERE credit_id = ?",
            (selected_credit_id,))
        amortization_data = c.fetchall()

        # Effacer les données actuelles du tableau d'amortissement
        self.amortization_table.delete(*self.amortization_table.get_children())

        # Remplir le tableau d'amortissement avec les nouvelles données
        for row in amortization_data:
            self.amortization_table.insert('', 'end', values=row)

        # Fermer la connexion à la base de données
        conn.close()
        self.update_entry_fields(selected_credit_id)

    def update_entry_fields(self, event=None):
        selected_credit = self.credit_dropdown.get()  # Récupérer le crédit sélectionné dans le dropdown
        if selected_credit:
            # Extraire le nom et prénom du client sélectionné
            nom_prenom = selected_credit.split(' - ')[0]

            # Remplacer les caractères de code client par une chaîne vide pour obtenir uniquement le nom et prénom
            nom_prenom = nom_prenom.replace(selected_credit.split(' - ')[-1], '').strip()

            # Extraire le credit_id du crédit sélectionné
            selected_credit_id = selected_credit.split(' - ')[-1]
            impayes = self.calcul_impaye_et_reste()[0]
            derniere_echance_np = self.calcul_impaye_et_reste()[2]

            # Activer les champs d'entrée pour permettre l'écriture
            self.entry_np.configure(state='normal')
            self.entry_cr.configure(state='normal')
            self.entry_cred.configure(state='normal')
            self.entry_reste.configure(state='normal')
            self.impaye_entry.configure(state='normal')
            self.date_np_entry.configure(state='normal')

            # Effacer les données actuelles des champs d'entrée
            self.entry_np.delete(0, 'end')
            self.entry_cr.delete(0, 'end')
            self.impaye_entry.delete(0, 'end')
            self.date_np_entry.delete(0, 'end')

            # Mettre à jour les champs d'entrée avec les nouvelles informations
            self.entry_np.insert(0, nom_prenom)  # Insérer le nom et prénom du client
            self.entry_cr.insert(0, selected_credit_id)  # Insérer le crédit ID
            # Modifier la taille de la police et la couleur du texte pour améliorer la lisibilité

            # Remplir automatiquement le montant du crédit et le reste du crédit
            montant_credit, reste_credit = self.fetch_credit_info(selected_credit_id)

            self.entry_cred.delete(0, 'end')
            self.entry_cred.insert(0, montant_credit)
            self.entry_cred.configure(state='disabled')

            self.entry_reste.delete(0, 'end')
            self.entry_reste.insert(0, reste_credit)
            self.impaye_entry.insert(0, f" {impayes} DT")
            self.entry_reste.configure(state='disabled')
            self.date_np_entry.insert(0, derniere_echance_np)
            self.date_np_entry.configure(state='disabled')
            # Désactiver à nouveau les champs d'entrée pour les rendre en lecture seule
            self.entry_np.configure(state='disabled')
            self.entry_cr.configure(state='disabled')
            self.entry_cred.configure(state='disabled')
            self.entry_reste.configure(state='disabled')
            self.impaye_entry.configure(state='disabled')
            numero_recu = self.get_last_receipt_number()

            self.entry_recu.delete(0, 'end')
            self.entry_recu.insert(0, numero_recu)
            self.entry_recu.configure(state='disabled')
            taux_recouvrement = self.taux_recouv_client()
            self.taux_recouv.configure(amountused=taux_recouvrement)

    def fetch_credit_info(self, credit_id):

        conn = sqlite3.connect('database.db')
        c = conn.cursor()

        # Récupérer le montant du crédit à partir de la table des crédits
        c.execute("SELECT montant FROM credits WHERE credit_id = ?", (credit_id,))
        montant_credit = c.fetchone()[0]

        # Récupérer le reste du crédit à partir de la table des paiements
        c.execute("SELECT Montant_reste FROM payements WHERE credit_id = ? ORDER BY reference_payement DESC LIMIT 1",
                  (credit_id,))
        dernier_payement = c.fetchone()
        if dernier_payement:
            reste_credit = dernier_payement[0]
        else:
            reste_credit = montant_credit  # Si aucun paiement n'a encore été effectué, le reste du crédit est le montant total

        conn.close()

        return montant_credit, reste_credit

    def get_last_receipt_number(self):
        # Connexion à la base de données
        conn = sqlite3.connect('database.db')
        cursor = conn.cursor()

        try:
            # Récupérer le dernier reçu depuis la base de données
            cursor.execute("SELECT Recu FROM payements ORDER BY ROWID DESC LIMIT 1")
            last_receipt = cursor.fetchone()

            # Extraire le numéro de reçu et l'année
            if last_receipt:
                last_receipt_parts = last_receipt[0].split('_')
                last_receipt_number = int(last_receipt_parts[0])
                last_receipt_year = int(last_receipt_parts[1])
            else:
                last_receipt_number = 0
                last_receipt_year = datetime.now().year

            # Vérifier si l'année a changé
            current_year = datetime.now().year
            if current_year != last_receipt_year:
                # Si l'année a changé, réinitialiser le numéro de reçu à 1
                new_receipt_number = 1
            else:
                # Sinon, incrémenter le numéro de reçu
                new_receipt_number = last_receipt_number + 1

            # Reformatter le numéro de reçu pour qu'il ait la même largeur que le numéro d'origine
            formatted_receipt_number = str(new_receipt_number).zfill(len(str(last_receipt_number)))

            # Construire la nouvelle chaîne de reçu
            new_receipt = f"{formatted_receipt_number}_{current_year}"

            return new_receipt

        except sqlite3.Error as e:
            print("Erreur lors de la récupération du dernier numéro de reçu:", e)
            return None
        finally:
            conn.close()

    def calcul_impaye_et_reste(self):

        selected_credit = self.credit_dropdown.get()  # Récupérer le crédit sélectionné dans le dropdown
        if selected_credit:
            # Extraire le credit_id du crédit sélectionné
            selected_credit_id = selected_credit.split(' - ')[-1]

        # Récupérer la date actuelle
        date_actuelle = datetime.now().date()

        # Initialiser les variables
        impayes = 0
        montant_reste = 0

        # Connexion à la base de données
        conn = sqlite3.connect('database.db')
        c = conn.cursor()

        try:
            # Calculer les impayés
            c.execute(
                "SELECT Montant_echeance FROM amortissement WHERE credit_id=? AND echeance_date <= ? AND paye = ?",
                (selected_credit_id, date_actuelle, ''))
            impayes_rows = c.fetchall()
            for row in impayes_rows:
                impayes += float(row[0].replace(' DT', ''))

            # Calculer le montant restant
            c.execute("SELECT montant FROM credits WHERE credit_id=?", (selected_credit_id,))
            montant_credit = float(str(c.fetchone()[0]).replace(' DT', ''))

            c.execute("SELECT SUM(Montant_payement) FROM payements WHERE credit_id=?", (selected_credit_id,))
            montant_paye = c.fetchone()[0]

            if montant_paye is None:
                montant_paye = 0
            else:
                montant_paye = float(montant_paye)

            montant_reste = montant_credit - montant_paye

            c.execute("SELECT min(echeance_date) FROM amortissement WHERE credit_id=? AND (paye=?or paye=?)",
                      (selected_credit_id, "", "P",))
            result = c.fetchone()[0]
            if result is not None:
                derniere_ech_np = result

        except sqlite3.Error as e:
            print("Erreur lors du calcul des impayés et du montant restant:", e)
        finally:
            conn.close()

        return impayes, montant_reste, derniere_ech_np

    def payement_partiel(self, credit_id=None):
        # Connexion à la base de données
        conn = sqlite3.connect('database.db')
        c = conn.cursor()
        c.execute("SELECT Payement_partiel FROM payements WHERE credit_id=? ORDER BY ROWID DESC ", (credit_id,))
        result = c.fetchone()
        c.close()

        # Vérifier si result est None et le remplacer par 0 si c'est le cas
        if result is None:
            partiel_paye = 0
        else:
            partiel_paye = result[0]

        print(partiel_paye)
        return partiel_paye

    def valider_montant_paiement(self, montant_paiement, reste_credit):
        # Vérifier si le montant de paiement est un float positif avec 3 chiffres après la virgule
        montant_paiement = float(montant_paiement)  # Convertir en float
        if not (montant_paiement > 0 and isinstance(montant_paiement, float)):
            messagebox.showerror("Erreur de montant",
                                 "Le montant de paiement doit être un float positif avec 3 chiffres après la virgule.")
            return False

        # Vérifier si le montant de paiement est supérieur au reste du crédit
        if montant_paiement > reste_credit:
            messagebox.showerror("Erreur de montant",
                                 "Le montant de paiement ne peut pas être supérieur au reste du crédit.")
            return False

        return True

    def effectuer_paiement(self):

        # Récupérer le montant de paiement depuis le champ entry_montant
        montant_paiement = self.entry_montant.get()

        # Récupérer la valeur restante du crédit depuis la fonction de calcul du reste du crédit
        impayes, montant_reste, derniere_ech_np = self.calcul_impaye_et_reste()

        # Valider le montant de paiement
        if not self.valider_montant_paiement(montant_paiement, montant_reste):
            # Si le montant de paiement n'est pas valide, ne pas procéder au paiement
            print("Le montant de paiement n'est pas valide. Veuillez vérifier et réessayer.")
            return
        # Récupérer le crédit sélectionné dans le dropdown
        selected_credit = self.credit_dropdown.get()
        if not selected_credit:
            print("Veuillez sélectionner un crédit.")
            return

        # Extraire le credit_id du crédit sélectionné
        selected_credit_id = selected_credit.split(' - ')[-1]

        # Récupérer le montant entré pour le paiement
        montant_paiement = float(self.entry_montant.get())

        # Récupérer la date du paiement
        date_paiement = self.date_pay.entry.get()
        partiel_paye = self.payement_partiel(selected_credit_id)

        # Vérifier si la date de paiement est valide
        try:
            date_format = '%d/%m/%Y'
            date_paiement = datetime.strptime(date_paiement, date_format).date()
        except ValueError:
            print("Format de date invalide. Utilisez le format DD-MM-YYYY.")
            return

        # Connexion à la base de données
        conn = sqlite3.connect('database.db')
        c = conn.cursor()

        try:
            # Récupérer toutes les traites non payées ou partiellement payées pour ce crédit
            c.execute(
                "SELECT echeance_numero, Montant_echeance FROM amortissement WHERE credit_id=? AND (paye = ? OR paye = ?)",
                (selected_credit_id, '', 'P'))
            traites_non_payees = c.fetchall()

            montant_restant = partiel_paye + montant_paiement
            for echeance_numero, montant_echeance in traites_non_payees:
                montant_echeance = float(montant_echeance.replace("DT", ""))  # Convertir en float
                montant_restant -= montant_echeance
                if montant_restant > 0:
                    # Le paiement couvre entièrement cette échéance, la marquer comme payée
                    c.execute(
                        "UPDATE amortissement SET date_paye=?, paye='X' WHERE credit_id=? AND echeance_numero=?",
                        (date_paiement, selected_credit_id, echeance_numero))
                elif montant_restant == 0:
                    c.execute(
                        "UPDATE amortissement SET date_paye=?, paye='X' WHERE credit_id=? AND echeance_numero=?",
                        (date_paiement, selected_credit_id, echeance_numero))
                    partiel_paye = 0
                    break
                else:
                    # Le paiement ne couvre pas entièrement cette échéance, la marquer comme partiellement payée
                    c.execute(
                        "UPDATE amortissement SET date_paye=?, paye='P' WHERE credit_id=? AND echeance_numero=?",
                        (date_paiement, selected_credit_id, echeance_numero))
                    # Sortir de la boucle car le paiement est partiel
                    partiel_paye = montant_echeance + montant_restant
                    break

            recu = self.get_last_receipt_number()
            mode_paye = self.modalite.get()

            # Insérer les données de paiement dans la table de paiements
            c.execute('''
                       INSERT INTO payements ( paye_date, Montant_payement, Montant_reste, Montant_impaye, Recu,Mode_paiement,Payement_partiel,credit_id)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                   ''',
                      (date_paiement, montant_paiement, max(montant_reste - montant_paiement, 0),
                       max(impayes - montant_paiement, 0), recu, mode_paye, partiel_paye, selected_credit_id))

            # Confirmer la transaction
            conn.commit()
            print("Paiement effectué avec succès.")
            self.update_entry_fields(selected_credit_id)
            self.update_amortization_table()

        except sqlite3.Error as e:
            conn.rollback()
            print("Erreur lors de l'effectuation du paiement:", e)
        finally:
            conn.close()

    def taux_recouv_client(self):
        date_du_jour = datetime.now().date()
        selected_credit = self.credit_dropdown.get()
        selected_credit_id = selected_credit.split(' - ')[-1]

        print(selected_credit)
        # Connexion à la base de données
        conn = sqlite3.connect('database.db')
        c = conn.cursor()
        if selected_credit:

            # Récupérer la somme des payements à partir de la table des crédits
            c.execute("SELECT SUM(Montant_payement) FROM payements WHERE credit_id=?", (selected_credit_id,))
            row = c.fetchone()[0]
            if row is not None:
                payement_total = int(row)
            else:
                payement_total = 0

            print(payement_total)
            # recuperer les échéance encourus
            try:
                # Récupérer toutes les traites non payées ou partiellement payées pour ce crédit
                c.execute(
                    "SELECT SUM( Montant_echeance),echeance_date FROM amortissement WHERE credit_id=? AND (echeance_date<=?)",
                    (selected_credit_id, date_du_jour,))
                row = c.fetchone()[0]
                if row is not None:
                    encourus = int(row)
                else:
                    encourus = 0

                print(encourus)

                if payement_total is not None and encourus > 0:
                    if encourus > payement_total:
                        taux_recouv = (payement_total / encourus) * 100
                    elif encourus < payement_total:
                        taux_recouv = 100
                    else:
                        taux_recouv = 0
                else:
                    # Gérer le cas où encourus est None
                    taux_recouv = 100
            except sqlite3.Error as e:
                conn.rollback()
                print("Erreur de calcul:", e)
            finally:
                c.close()
        return (round(taux_recouv))

    def taux_recouvrement_global(self,mois_var,annee_var):
        taux_recouv_glob = 0
        mois = mois_var.get().lower()
        annee = int(annee_var.get())

        mois = int(mois.get().split('/')[1])  # Supposons que la date soit au format 'dd/mm/yyyy'

        # Obtenir le dernier jour du mois pour l'année et le mois donnés
        dernier_jour = calendar.monthrange(annee, mois)[1]

        # Créer la date maximale
        date_maximale = f"{annee}-{mois}-{dernier_jour}"
        print(date_maximale)

        # Connexion à la base de données
        conn = sqlite3.connect('database.db')
        c = conn.cursor()
        c.execute(
            "SELECT SUM(Montant_payement) FROM payements WHERE paye_date <=?)",
            ( dernier_jour,))
        tot_payement = c.fetchone()[0]
        c.execute(
            "SELECT SUM( Montant_echeance) FROM amortissement WHERE   (echeance_date<=?)",
            ( dernier_jour,))
        tot_echeance = c.fetchone()[0]
        taux_recouv_glob = round((tot_payement/tot_echeance),0)* 100
        print(f"le taux de recouvrement total est de : {taux_recouv} %")
        return taux_recouv





    def mettre_a_jour_taux_recouvrement(self, taux_recouvrement=100):
        taux_recouvrement = self.taux_recouv_client()
        self.taux_recouv.Meter.config(amountused=taux_recouvrement)

    def adapt_datetime(ts):
        return ts.strftime("%Y-%m-%d ")

    def convert_datetime(s):
        return datetime.datetime.strptime(s, "%Y-%m-%d ")

        ### CLASSE IMPRESSION / TOUS CE QUI CONCERNE IMPRESSION RECU CONTRAT ECT


class ImpressionManager(tbk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.application = kwargs.get('application')
        self.credits_section = kwargs.get('credits_section')
        self.imf_section = kwargs.get('imf_section')
        self.fichier_temporaire = None
        pdf_counter = 1

    def choisir_imprimante(self, imprimantes_var, nom_fichier):
        self.fenetre = Tk()  # Créez la fenêtre ici
        self.fenetre.title("Choisir une imprimante")
        self.fenetre.geometry('300x200+400+100')  # Ajustez la géométrie pour une meilleure apparence

        def imprimer():
            self.imprimante_choisie = self.imprimantes_var.get()
            print(f"Imprimante choisie : {self.imprimante_choisie}")

            if self.imprimante_choisie:
                # Utilisez ctypes.windll.shell32 pour imprimer avec l'imprimante choisie
                print(f"Imprimer avec : {self.imprimante_choisie}")
                ctypes.windll.shell32.ShellExecuteW(None, "print", nom_fichier, None, None, 0)
            else:
                print("Aucune imprimante choisie.")

            # Détruire la fenêtre après l'impression
            self.fenetre.destroy()

        Label(self.fenetre, text="Choisissez une imprimante:").pack(pady=10)

        # Utilisez ttk.Combobox pour une meilleure gestion de la taille
        self.imprimantes = [printer['pPrinterName'] for printer in
                            win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS,
                                                    None, 5)]

        combobox = tbk.Combobox(self.fenetre, textvariable=imprimantes_var, values=self.imprimantes)
        combobox.pack(pady=10)

        tbk.Button(self.fenetre, text="Imprimer", command=imprimer).pack(pady=10)

        # Exécute la boucle principale de la fenêtre
        self.fenetre.mainloop()

        # Retourne True si une imprimante a été choisie, False sinon
        return bool(self.imprimante_choisie)

    def imprimer_pdf(self, packet=None, nom_fichier=""):
        self.root = Tk()
        self.root.withdraw()  # Cacher la fenêtre principale

        # Utilisez win32print.EnumPrinters pour obtenir la liste des imprimantes installées
        self.imprimantes = [printer['pPrinterName'] for printer in
                            win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 5)]

        self.imprimantes_var = StringVar(self.root)
        self.imprimantes_var.set(self.imprimantes[0])

        # Utilisez une boîte de dialogue personnalisée pour choisir une imprimante
        confirmation_impression = self.choisir_imprimante(self.imprimantes_var, nom_fichier)

        # Détruire la fenêtre principale après la sélection de l'imprimante
        self.root.destroy()

        # Si l'utilisateur a confirmé l'impression
        if confirmation_impression:
            # Demander la confirmation avant l'impression
            confirmation = messagebox.askokcancel("Confirmation", "Voulez-vous vraiment imprimer le document ?")

            if confirmation:
                # Imprimer le document
                ctypes.windll.shell32.ShellExecuteW(None, "print", nom_fichier, None, None, 0)
                # Afficher un message après l'impression
                messagebox.showinfo("Impression réussie", "Le document a été imprimé avec succès.")
            else:
                # L'utilisateur a annulé l'impression
                messagebox.showinfo("Impression annulée", "L'impression a été annulée par l'utilisateur.")
        else:
            # L'utilisateur a annulé la sélection de l'imprimante
            messagebox.showinfo("Impression annulée",
                                "L'impression a été annulée car aucune imprimante n'a été sélectionnée.")

    def generate_and_print_table(self, Nom, Prenom, CIN):
        conn = sqlite3.connect('database.db')
        c = conn.cursor()

        try:
            # Recherche du client_id correspondant aux valeurs de Nom, Prenom, et CIN dans la table clients
            c.execute('''
                          SELECT client_id FROM clients 
                          WHERE Nom = ? AND Prenom = ? AND CIN = ?
                      ''', (Nom, Prenom, CIN))

            result_1 = c.fetchone()
            if result_1:
                selected_client_id = result_1[0]

                # Obtenez le dernier credit_id pour le client sélectionné
                c.execute('''
                                 SELECT credit_id
                                 FROM credits
                                 WHERE client_id = ?
                                 ORDER BY ROWID DESC
                                 LIMIT 1
                             ''', (selected_client_id,))

                selected_credit_id = c.fetchone()[0]

                if selected_credit_id:
                    # Obtenez les informations du dernier crédit
                    c.execute('''
                                 SELECT nom, prenom, cin, montant, duree
                                 FROM credits
                                 WHERE credit_id = ?
                             ''', (selected_credit_id,))

                    credit_info = c.fetchone()

                    # Obtenez les données d'amortissement ordonnées par la première échéance pour le dernier crédit
                    c.execute('''
                                    SELECT echeance_date, echeance_numero, Montant_echeance,Interet,Reste_du_credit
                                    FROM amortissement
                                    WHERE credit_id = ?
                                    ORDER BY echeance_numero
                                ''', (selected_credit_id,))

                    amortization_data = c.fetchall()
                    print(amortization_data)

                    if credit_info:
                        nom_client, prenom_client, cin_client, montant_credit, duree_credit = credit_info

                    else:
                        nom_client, prenom_client, cin_client, montant_credit, duree_credit = (
                            None, None, None, None, None)

                    # Afficher une boîte de dialogue demandant à l'utilisateur s'il veut imprimer ou sauvegarder
                    user_choice = messagebox.askquestion("Choix",
                                                         "Voulez-vous imprimer ou sauvegarder le tableau d'amortissement ?")

                    if user_choice == 'yes':

                        pdfmetrics.registerFont(TTFont('SimplifiedArabic', "Document/NotoSansArabic-Regular.ttf"))
                        # Création du document PDF
                        pdf_filename = "tableau_amortissement.pdf"
                        pdf = SimpleDocTemplate(pdf_filename, pagesize=A4)

                        elements = []

                        # Ajouter un en-tête centré
                        en_tete = "جدول اللاستخلاصات"
                        en_tete = get_display(arabic_reshaper.reshape(en_tete))
                        style = ParagraphStyle(name='CenteredStyle', alignment=1, fontSize=12,
                                               fontName="SimplifiedArabic")
                        elements.append(Paragraph(en_tete, style))
                        elements.append(Spacer(1, 20))
                        listes = [nom_client, prenom_client]

                        # styles = str(getSampleStyleSheet())[::-1]
                        arabic_list = [get_display(arabic_reshaper.reshape(text)) for text in listes]

                        print(arabic_list)
                        # Ajouter les informations client à gauche
                        info_client = f"<b> </b> {arabic_list[1]}<br/><b>  </b>{arabic_list[0]}<br/><b></b> {cin_client}<br/><b> </b>   {montant_credit} "
                        style = ParagraphStyle(name='RTLStyle', alignment=TA_RIGHT, fontSize=12,
                                               fontName="SimplifiedArabic")
                        elements.append(Paragraph(info_client, style))
                        elements.append(Spacer(1, 24))
                        elements.append(Spacer(1, 24))
                        titre_tableaux = ["الباقي", "الفوائض ", "المبلغ ", "رقم", "تاريخ الإستخلاص"]
                        titre_arabe = [get_display(arabic_reshaper.reshape(texte)) for texte in titre_tableaux]

                        # Ajouter le tableau d'amortissement
                        data = [[titre_arabe[4], titre_arabe[3], titre_arabe[2], titre_arabe[1],
                                 titre_arabe[0]]] + amortization_data
                        tableau = Table(data, colWidths=[100, 100, 100], rowHeights=25)
                        style = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                            ('FONTNAME', (0, 0), (-1, 0), 'SimplifiedArabic'),
                                            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                                            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                                            ('GRID', (0, 0), (-1, -1), 1, colors.black)])
                        tableau.setStyle(style)
                        elements.append(tableau)
                        elements.append(Spacer(1, 2))

                        # Ajouter un pied de page (IMF Data)
                        try:
                            with open("donnees_imf.json", "r") as json_file:
                                imf_data = json.load(json_file)
                        except FileNotFoundError:
                            messagebox.showerror("Erreur", "Fichier JSON introuvable.")
                            return
                        except json.decoder.JSONDecodeError:
                            messagebox.showerror("Erreur", "Erreur de décodage JSON.")
                            return

                        imf_style = TableStyle([
                            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                            ('FONTNAME', (0, 0), (-1, -1), 'SimplifiedArabic'),
                            ('FONTSIZE', (0, 0), (-1, -1), 7),
                            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                            ('LEADING', (0, 0), (-1, -1), 1),
                            ('ALIGN', (0, 0), (0, -1), 'RIGHT'),
                            ('LEFTPADDING', (5, 2), (5, 2), 20),
                            ('ALIGN', (4, 2), (4, 2), 'RIGHT'),
                        ])

                        imf_table_data = [
                            [imf_data.get('Tel'), imf_data['Adresse'].get('Compte 2'),
                             imf_data['Adresse'].get('Compte 1'), imf_data.get('Raison Sociale')],
                            [imf_data['RIB'].get('Compte 1'), imf_data['Banque'].get('Compte 1')],
                            [imf_data.get('CNSS'), imf_data.get('RNE'), imf_data.get('MAIL')]]
                        imf_table_data_arabe = [
                            [get_display(arabic_reshaper.reshape(text)) for text in row]
                            # Appliquer reshape à chaque élément de la ligne
                            for row in imf_table_data]

                        imf_table = Table(imf_table_data_arabe, colWidths=[160, 160, 80, 80, 80, 80, 80], rowHeights=15)
                        imf_table.setStyle(imf_style)
                        imf_style.add('RIGHTPADDING', (0, 0), (0, -1), 10)
                        elements.append(imf_table)

                        # Construire le PDF avec les éléments
                        pdf.build(elements)
                        self.imprimer_pdf(nom_fichier=pdf_filename)

                        # Afficher une boîte de dialogue pour informer l'utilisateur que le PDF a été imprimé
                        messagebox.showinfo("Information", "Le tableau d'amortissement a été imprimé.")

                    else:
                        # Sauvegarder le tableau d'amortissement avec un nom de fichier unique
                        timestamp = time.strftime("%Y%m%d")
                        saved_filename = f"tableau_amortissement_{timestamp}.pdf"
                        pdf.build(elements, filename=saved_filename)

                        # Afficher une boîte de dialogue pour informer l'utilisateur que le tableau a été sauvegardé
                        messagebox.showinfo("Information", "Le tableau d'amortissement a été sauvegardé.")

            else:
                # Afficher un message d'erreur si le client n'est pas trouvé
                messagebox.showerror("Erreur", "Client non trouvé dans la base de données.")

        finally:
            conn.close()

    def remplir_traite(self, can, nom_client, prenom_client, cin_client, echeance_numero, echeance_date,
                       montant_echeance, traite_num):
        pdfmetrics.registerFont(TTFont('SimplifiedArabic', "Document/NotoNaskhArabic-Regular.ttf"))
        pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))

        # Ajustez les positions en fonction du numéro de traite dans le groupe de trois
        y_position = 710 - (traite_num - 1) * 290
        x_positions = [220, 350, 400, 470, 230, 320, 450]
        listes = [nom_client, prenom_client]
        for text in listes:
            styles = str(getSampleStyleSheet())[::-1]
            arabic_list = [arabic_reshaper.reshape(text) for text in listes]

        can.setFont('SimplifiedArabic', 15)
        can.drawString(x_positions[1], y_position, get_display(arabic_list[0]))
        can.drawString(x_positions[2], y_position, get_display(arabic_list[1]))
        can.setFont('Arial', 15)
        can.drawString(x_positions[2], y_position - 40, str(cin_client))
        can.drawString(x_positions[3], y_position - 110, str(echeance_numero))
        can.drawString(x_positions[4], y_position + 120, str(echeance_date))
        can.drawString(x_positions[5], y_position + 50, str(montant_echeance))
        can.drawString(x_positions[6], y_position + 120, str(montant_echeance))

    def generate_and_print_traites(self, Nom, Prenom, CIN):
        conn = sqlite3.connect('database.db')
        c = conn.cursor()
        try:
            # Recherche du client_id correspondant aux valeurs de Nom, Prenom, et CIN dans la table clients
            c.execute('''
                SELECT client_id FROM clients 
                WHERE Nom = ? AND Prenom = ? AND CIN = ?
                ORDER BY client_id DESC
                LIMIT 1
            ''', (Nom, Prenom, CIN))

            result_1 = c.fetchone()
            if result_1:
                selected_client_id = result_1[0]

                # Obtenez le dernier credit_id pour le client sélectionné
                c.execute('''
                    SELECT credit_id
                    FROM credits
                    WHERE client_id = ?
                    ORDER BY ROWID DESC
                    LIMIT 1
                ''', (selected_client_id,))

                selected_credit_id = c.fetchone()
                if selected_credit_id:
                    selected_credit_id = selected_credit_id[0]  # Accès à la première colonne du résultat

                    # Obtenez les données d'amortissement pour le selected_credit_id
                    c.execute('''
                        SELECT echeance_date, echeance_numero, Montant_echeance, Interet, Reste_du_credit
                        FROM amortissement
                        WHERE credit_id = ?
                        ORDER BY echeance_numero
                    ''', (selected_credit_id,))

                    amortization_data = c.fetchall()

                    # Chargez le modèle PDF existant
                    pdf_template = PdfReader("Document/traites.pdf")

                    # Créez un nouveau PDF pour la sortie
                    output_pdf = PdfWriter()

                    # Nombre total de traites
                    total_traites = len(amortization_data)

                    # Calcul du nombre de pages nécessaire
                    pages_needed = total_traites // 3
                    if total_traites % 3 != 0:
                        pages_needed += 1

                    # Création d'une liste de traites par page
                    traites_per_page = [amortization_data[i:i + 3] for i in range(0, total_traites, 3)]

                    # Itération sur les pages
                    for page_num in range(pages_needed):
                        # Création d'un nouveau document PDF temporaire pour chaque page
                        packet = BytesIO()
                        c = canvas.Canvas(packet, pagesize=A4)

                        # Itération sur les traites sur cette page
                        for traite_num, (
                                echeance_date, echeance_numero, montant_echeance, interet,
                                reste_du_credit) in enumerate(
                            traites_per_page[page_num], start=1):
                            # Appel de la fonction pour remplir les données de la traite
                            self.remplir_traite(c, Nom, Prenom, CIN, echeance_numero, echeance_date,
                                                montant_echeance, traite_num)

                        # Sauvegarde du document PDF temporaire
                        c.save()

                        # Déplacez le curseur de l'objet BytesIO au début du flux
                        packet.seek(0)

                        # Chargez le modèle PDF initial pour chaque nouvelle page
                        pdf_template_copy = PdfReader("Document/traites.pdf")

                        # Chargez le canevas modifié en tant que document PDF
                        new_pdf = PdfReader(packet)

                        # Fusionnez chaque page du modèle PDF initial avec le canevas modifié
                        for page in pdf_template_copy.pages:
                            page.merge_page(new_pdf.pages[0])
                            output_pdf.add_page(page)

                    # Sauvegarde du document PDF final
                    with open("traites_final.pdf", "wb") as output_file:
                        output_pdf.write(output_file)

                conn.close()
                messagebox.showinfo("Succès", "Traités générés avec succès!")
        except Exception as e:
            # Gérer les erreurs (afficher un message, journaliser, etc.)
            print(f"Erreur lors de la génération et de l'impression des traites : {e}")
            messagebox.showerror("Erreur", f"Erreur lors de la génération et de l'impression des traites : {e}")
        finally:
            self.imprimer_pdf(nom_fichier="C:/Users/User/PycharmProjects/ligne_credits/traites_final.pdf")
            print("Connexion fermée avec succès")

    def remplir_contrat(self, can, page, nom_client, prenom_client, cin_client, date_delivrance, adresse, delegation,
                        secteur_activite, sous_secteur, date_credit, montant, duree, raison_sociale, adresse_imf,
                        date_premiere_echeance, date_derniere_echeance):

        pdfmetrics.registerFont(TTFont('Arabic', "C:\\apps\\29ltbukraregular.ttf"))
        pdfmetrics.registerFont(TTFont('SimplifiedArabic', "Document/NotoNaskhArabic-Regular.ttf"))
        pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
        x_positions = [320, 410, 450, 150, 130, 350, 400, 100, 480, 300]  # Ajoutez ici les positions ajustées

        listes = [str(nom_client), str(prenom_client), str(adresse), str(delegation), str(secteur_activite),
                  str(sous_secteur), str(raison_sociale),
                  str(adresse_imf)]

        for text in listes:
            styles = str(getSampleStyleSheet())[::-1]
            arabic_list = [arabic_reshaper.reshape(text) for text in
                           listes]
            can.setFont('SimplifiedArabic', 15)

            if page == 0:
                can.drawString(x_positions[5], 280, get_display(arabic_list[6]))
                can.drawString(x_positions[6], 250, get_display(arabic_list[7]))
                can.drawString(x_positions[0], 170, get_display(arabic_list[0]))
                can.drawString(x_positions[1], 170, get_display(arabic_list[1]))
                can.drawString(x_positions[2], 130, str(cin_client))
                can.drawString(x_positions[2], 100, get_display(arabic_list[2]))
                can.drawString(240, 100, get_display(arabic_list[3]))

                can.setFont('Arial', 15)
                can.drawString(x_positions[3], 130, str(date_delivrance))

            elif page == 1:
                can.setFont('SimplifiedArabic', 15)
                # Ajoutez ici les positions ajustées pour la page 1
                can.drawString(x_positions[5], 730, get_display(arabic_list[6]))
                can.drawString(x_positions[3], 730, get_display(arabic_list[1]))
                can.drawString(x_positions[7], 730, get_display(arabic_list[0]))
                can.drawString(x_positions[5], 710, str(montant))
                can.drawString(450, 690, get_display(arabic_list[5]))
                can.drawString(280, 690, get_display(arabic_list[4]))
                can.drawString(x_positions[7], 690, str(duree))
                can.drawString(x_positions[7], 540, str(montant))
                can.drawString(x_positions[8], 510, str(duree))
                can.drawString(x_positions[6], 300, str(duree))
                can.setFont('Arial', 15)
                can.drawString(x_positions[9], 510, str(date_premiere_echeance))
                can.drawString(x_positions[4], 510, str(date_derniere_echeance))

            else:
                can.setFont('Arial', 15)
                can.drawString(x_positions[7], 410, date_credit)
                can.setFont('SimplifiedArabic', 15)
                can.drawString(290, 410, get_display(arabic_list[3]))
                can.drawString(x_positions[2], 200, get_display(arabic_list[0]))
                can.drawString(500, 200, get_display(arabic_list[1]))

    def generate_and_print_contrat(self, Nom, Prenom, CIN):
        conn = sqlite3.connect('database.db')
        c = conn.cursor()

        Date_premiere_echeance = None
        Date_derniere_echeance = None

        try:
            # Recherche du client_id correspondant aux valeurs de Nom, Prenom et CIN dans la table clients
            c.execute('''
                   SELECT client_id, date_delivrance_cin, Date_naissance, adresse, delegation, secteur_activite, sous_secteur
                   FROM clients 
                   WHERE Nom = ? AND Prenom = ? AND CIN = ?
               ''', (Nom, Prenom, CIN))

            result_client = c.fetchone()

            if result_client:
                selected_client_id, date_delivrance_cin, Date_naissance, adresse, delegation, secteur_activite, sous_secteur = result_client

                # Obtenez le dernier credit_id pour le client sélectionné
                c.execute('''
                       SELECT credit_id
                       FROM credits
                       WHERE client_id = ?
                       ORDER BY ROWID DESC
                       LIMIT 1
                   ''', (selected_client_id,))

                selected_credit_id = c.fetchone()[0]

                if selected_credit_id:
                    # Obtenez les informations du dernier crédit
                    c.execute('''
                           SELECT Nom, Prenom, CIN, Date_Credit, montant, duree
                           FROM credits
                           WHERE credit_id = ?
                       ''', (selected_credit_id,))

                    credit_info = c.fetchone()

                    # Obtenez les données d'amortissement ordonnées par la première échéance pour le dernier crédit
                    c.execute('''
                           SELECT MIN(echeance_date) AS premiere_echeance, MAX(echeance_date) AS derniere_echeance
                           FROM amortissement
                           WHERE credit_id = ?
                       ''', (selected_credit_id,))

                    dates_amortissement = c.fetchone()

                    if credit_info:
                        nom_client, prenom_client, cin_client, date_credit, montant_credit, duree_credit = credit_info
                    else:
                        nom_client, prenom_client, cin_client, date_credit, montant_credit, duree_credit = (
                            None, None, None, None, None, None)

                    if dates_amortissement:
                        Date_premiere_echeance, Date_derniere_echeance = dates_amortissement
                    else:
                        Date_premiere_echeance, Date_derniere_echeance = (None, None)

                    # Chargez les informations IMF à partir du fichier JSON
                    try:
                        with open("donnees_imf.json", "r") as json_file:
                            imf_data = json.load(json_file)
                            raison_sociale = imf_data.get('Raison Sociale', '')
                            adresse_imf = imf_data['Adresse'].get('Compte 1', '')
                    except FileNotFoundError:
                        raison_sociale, adresse_imf = (None, None)
                        messagebox.showerror("Erreur", "Fichier JSON introuvable.")
                    except json.decoder.JSONDecodeError:
                        raison_sociale, adresse_imf = (None, None)
                        messagebox.showerror("Erreur", "Erreur de décodage JSON.")
                        # Chargez le modèle PDF existant
                    print("Chargement du modèle PDF existant...")
                    pdf_template = PdfReader("Document/contrat_credit.pdf", "rb")

                    # Créez un nouveau PDF pour la sortie
                    output_pdf = PdfWriter()

                    # Parcourez les pages du modèle PDF initial avant d'insérer le remplissage du contrat
                    for template_page_num in range(len(pdf_template.pages)):
                        # Créez un nouveau document PDF temporaire
                        packet = BytesIO()
                        can = canvas.Canvas(packet, pagesize=letter)

                        # Reste du code pour générer le contrat
                        print("Remplissage du contrat...")
                        self.remplir_contrat(can, template_page_num, nom_client, prenom_client, cin_client,
                                             date_delivrance_cin, adresse, delegation, secteur_activite,
                                             sous_secteur,
                                             date_credit, montant_credit, duree_credit, raison_sociale, adresse_imf,
                                             Date_premiere_echeance, Date_derniere_echeance)

                        # Sauvegarde du document PDF temporaire
                        print("Sauvegarde du document PDF temporaire...")
                        can.save()

                        # Déplacez le curseur de l'objet BytesIO au début du flux
                        packet.seek(0)

                        # Chargez le canevas modifié en tant que document PDF
                        new_pdf = PdfReader(packet)

                        # Récupérez la page du modèle PDF initial
                        template_page = pdf_template.pages[template_page_num]

                        # Fusionnez le contenu de la page remplie avec la page du modèle
                        template_page.merge_page(new_pdf.pages[0])

                        # Ajoutez la page remplie au document PDF final
                        output_pdf.add_page(template_page)

                    # Enregistrez le document PDF final
                    temp_filename = "C:/Users/User/PycharmProjects/pythonProject4/temp_contract.pdf"
                    with open(temp_filename, "wb") as temp_file:
                        output_pdf.write(temp_file)

                    # Imprimez le document PDF final
                    print("Impression du document PDF final...")
                    self.imprimer_pdf(nom_fichier=temp_filename)
                    messagebox.showinfo("Succès", "Impression du contrat avec succee!")



        except Exception as e:
            # Gérer les erreurs (afficher un message, journaliser, etc.)
            messagebox.showwarning("Attention","Erreur lors de l'impression !")
            print(f"Erreur lors de la génération et de l'impression du contrat : {e}")

        finally:
            # Fermeture de la connexion uniquement si le bloc try a réussi
            conn.close()
            print("Connexion fermée avec succès")

    def generer_rapport_journalier(self):
        # Créer une fenêtre pop-up pour saisir les dates de début et de fin
        popup = Toplevel(self.application)
        popup.title("Générer rapport journalier")
        popup.geometry("800x600")

        # Labels et champs de saisie pour les dates de début et de fin
        Label(popup, text="Date de début :").pack(pady=(20, 5))  # Laisser 20 pixels au-dessus et 5 pixels en dessous
        debut_entry = tbk.DateEntry(popup)
        debut_entry.pack(pady=(20, 20))

        Label(popup, text="Date de fin :").pack(pady=(5, 20))  # Laisser 5 pixels au-dessus et 20 pixels en dessous
        fin_entry = tbk.DateEntry(popup)
        fin_entry.pack(pady=(20, 5))

        # Ajouter un espace vide entre le bouton et le bas de la fenêtre
        bottom_frame = Frame(popup)
        bottom_frame.pack(side="bottom", pady=20)
        # Bouton pour générer le rapport
        tbk.Button(bottom_frame, text="Générer rapport", bootstyle="success",
                   command=lambda: self.generer_rapport(popup, debut_entry.entry.get(), fin_entry.entry.get())).pack()

    def generer_rapport(self, popup, debut, fin,total_paiements=0, montant_total=0):
        temp_info_list=[]
        # Convertir les dates de début et de fin au format datetime
        date_format = '%d/%m/%Y'
        date_debut = datetime.strptime(debut, date_format).date()
        date_fin = datetime.strptime(fin, date_format).date()

        # Création de la fenêtre principale
        rapport_window = Toplevel(popup)
        rapport_window.title("Rapport journalier")
        rapport_window.geometry("1200x800")

        # Création de la grille
        rapport_window.grid_rowconfigure(0, weight=1)
        rapport_window.grid_rowconfigure(1, weight=1)
        rapport_window.grid_rowconfigure(2, weight=1)
        rapport_window.grid_columnconfigure(0, weight=1)
        rapport_window.grid_columnconfigure(1, weight=1)

        # Titre du rapport centré dans le tableau
        titre_label = Label(rapport_window, text=f"Recouvrement journalier de {debut} à {fin}",
                            font=('Calibri', 14, 'bold'))
        titre_label.grid(row=0, column=1, columnspan=1, pady=(20, 0), sticky="ew")

        style = tbk.Style()
        style.configure("mystyle.Treeview", highlightthickness=1, bd=0,
                        font=('Calibri', 11))  # Modify the font of the body
        style.configure("mystyle.Treeview.Heading", font=('Calibri', 12, 'bold'), background="#808080")
        # Création du Treeview

        tree = tbk.Treeview(rapport_window, style="mystyle.Treeview")

        tree["columns"] = ("Nom", "Prénom", "CIN", "Date", "Montant", "Recu", "Mode de pai")
        # Réduire l'espace entre les colonnes
        tree.column("#0", width=100)
        for col in tree["columns"]:
            tree.column(col, width=100)

        # Aligner les titres des colonnes avec les données
        tree.heading("#0", text="Crédit ID", anchor="w")
        for col in tree["columns"]:
            tree.heading(col, text=col, anchor="w")

        tree.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=20)

        # Récupérer les crédits dans la plage de dates spécifiée
        conn = sqlite3.connect("database.db")
        cursor = conn.cursor()
        cursor.execute("SELECT credit_id FROM payements WHERE paye_date BETWEEN ? AND ?", (date_debut, date_fin))
        credit_ids = [row[0] for row in cursor.fetchall()]
        # Parcourir les paiements et récupérer les informations des clients correspondants
        total_paiements = 0
        montant_total = 0

        # Parcourir les crédits et récupérer les informations des clients correspondants
        for credit_id in credit_ids:
            client_id = self.recuperer_client_id(credit_id)[0]
            client_info = self.recuperer_donnees_clients(client_id)
            paiements_info = self.recuperer_informations_paiement(credit_id)

            # Afficher les informations dans le Treeview
            for paiement in paiements_info:
                tree.insert("", "end", text=credit_id, values=(client_info[0], client_info[1], client_info[2],
                                                               paiement[0], paiement[1], paiement[2], paiement[3]))
                temp_info_list.append((credit_id,) + client_info + paiement)
                total_paiements += 1
                montant_total += paiement[1]

        # Ajouter deux lignes vides
        tree.insert("", "end", values=("", "", "", "", "", "", "", ""))
        tree.insert("", "end", values=("", "", "", "", "", "", "", ""))

        # Ajouter une ligne pour le total des paiements
        tree.insert("", "end", values=("Total", total_paiements, "", "", f" {montant_total} DT", "", "", ""))
        conn.close()

        # Ajouter un espace vide en bas
        espace_vide_bas = Frame(rapport_window, height=20)
        espace_vide_bas.grid(row=2, column=0, columnspan=2, pady=(0, 20), sticky="nsew")

        # Ajouter un espace vide entre le bouton et le bas de la fenêtre
        bottom_frame = Frame(rapport_window)
        bottom_frame.grid(row=3, column=0, columnspan=2, pady=(20, 0))

        # Bouton d'impression au centre
        Button(bottom_frame, text="Imprimer", command=lambda: self.generer_rapport_journalier_pdf(temp_info_list,date_debut,date_fin,total_paiements,montant_total)).pack(pady=10)

        print(temp_info_list,total_paiements,montant_total)
        return (temp_info_list,total_paiements,montant_total)
    def recuperer_donnees_clients(self, client_id):
        conn = sqlite3.connect("database.db")
        cursor = conn.cursor()

        # Sélectionner les informations du client à partir du credit_id
        cursor.execute("SELECT nom, prenom, cin FROM clients WHERE client_id = ?", (client_id,))
        row = cursor.fetchone()
        conn.close()

        return row

    # Fonction pour récupérer les informations de paiement
    def recuperer_informations_paiement(self, credit_id):
        conn = sqlite3.connect("database.db")
        cursor = conn.cursor()

        # Sélectionner les informations de paiement à partir du credit_id
        cursor.execute("SELECT paye_date, Montant_payement, Recu, Mode_paiement FROM payements WHERE credit_id = ?",
                       (credit_id,))
        rows = cursor.fetchall()
        conn.close()

        return rows
        # Fonction pour générer le rapport de paiement

    def recuperer_client_id(self, credit_id):
        conn = sqlite3.connect("database.db")
        cursor = conn.cursor()
        cursor.execute("SELECT client_id FROM credits WHERE credit_id = ?", (credit_id,))
        client_id = cursor.fetchone()
        conn.close()
        return client_id

    def generer_rapport_mensuel(self):
        # Créer une fenêtre pop-up pour saisir les dates de début et de fin
        popup = Toplevel(self.application)
        popup.title("Générer rapport Mensuel")
        popup.geometry("800x600")

        # Labels et champs de saisie pour sélectionner le mois et l'année
        Label(popup, text="Mois :").pack(pady=5)
        self.mois_var = StringVar()
        self.mois_combo = tbk.Combobox(popup, textvariable=self.mois_var, values=self.get_months())
        self.mois_combo.pack()

        Label(popup, text="Année :").pack(pady=5)
        self.annee_var = StringVar()
        self.annee_combo = tbk.Combobox(popup, textvariable=self.annee_var, values=self.get_years())
        self.annee_combo.pack()
        # Ajouter un espace vide entre le bouton et le bas de la fenêtre
        bottom_frame = Frame(popup)
        bottom_frame.pack(side="bottom", pady=20)
        # Bouton pour générer le rapport
        tbk.Button(bottom_frame, text="Générer rapport", bootstyle="success",
                   command=lambda: self.generer_rapport_mens(popup, self.mois_var, self.annee_var)).pack()

    def get_months(self):
        return ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre',
                'Novembre', 'Décembre']

    def get_years(self):
        current_year = datetime.now().year
        return [str(year) for year in range(2023, current_year + 1)]

    def generer_rapport_mens(self, popup, mois_var, annee_var):
        temp_info_list=[]
        # Extraire les valeurs sélectionnées des objets StringVar
        mois = mois_var.get().lower()
        annee = int(annee_var.get())

        # Convertir le nom du mois en numéro de mois
        mois_numerique = {mois: idx for idx, mois in enumerate(calendar.month_name) if mois}
        print(mois_numerique)
        mois_num = mois_numerique.get(mois)
        premier_jour = datetime(annee, mois_num, 1)
        dernier_jour = datetime(annee, mois_num % 12 + 1, 1) - timedelta(days=1) if mois_num != 12 else datetime(annee,
                                                                                                                 12, 31)

        # Création de la fenêtre principale
        rapport_window = Toplevel(popup)
        rapport_window.title("Rapport Mensuel")
        rapport_window.geometry("1200x800")

        # Création de la grille
        rapport_window.grid_rowconfigure(0, weight=1)
        rapport_window.grid_rowconfigure(1, weight=1)
        rapport_window.grid_rowconfigure(2, weight=1)
        rapport_window.grid_columnconfigure(0, weight=1)
        rapport_window.grid_columnconfigure(1, weight=1)

        # Titre du rapport centré dans le tableau
        titre_label = Label(rapport_window, text=f"Recouvrement mensuel du mois  de {mois} {annee}",
                            font=('Calibri', 14, 'bold'))
        titre_label.grid(row=0, column=1, columnspan=1, pady=(20, 0), sticky="ew")

        style = tbk.Style()
        style.configure("mystyle.Treeview", highlightthickness=1, bd=0,
                        font=('Calibri', 11))  # Modify the font of the body
        style.configure("mystyle.Treeview.Heading", font=('Calibri', 12, 'bold'), background="#808080")
        # Création du Treeview

        tree = tbk.Treeview(rapport_window, style="mystyle.Treeview")

        tree["columns"] = ("Nom", "Prénom", "CIN", "Date", "Montant", "Recu", "Mode de pai")
        # Réduire l'espace entre les colonnes
        tree.column("#0", width=100)
        for col in tree["columns"]:
            tree.column(col, width=100)

        # Aligner les titres des colonnes avec les données
        tree.heading("#0", text="Crédit ID", anchor="w")
        for col in tree["columns"]:
            tree.heading(col, text=col, anchor="w")

        tree.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=20)

        # Récupérer les crédits dans la plage de dates spécifiée
        conn = sqlite3.connect("database.db")
        cursor = conn.cursor()
        cursor.execute("SELECT credit_id FROM payements WHERE paye_date BETWEEN ? AND ?", (premier_jour, dernier_jour))
        credit_ids = [row[0] for row in cursor.fetchall()]
        # Parcourir les paiements et récupérer les informations des clients correspondants
        total_paiements = 0
        montant_total = 0

        # Parcourir les crédits et récupérer les informations des clients correspondants
        for credit_id in credit_ids:
            client_id = self.recuperer_client_id(credit_id)[0]
            client_info = self.recuperer_donnees_clients(client_id)
            paiements_info = self.recuperer_informations_paiement(credit_id)

            # Afficher les informations dans le Treeview
            for paiement in paiements_info:
                tree.insert("", "end", text=credit_id, values=(client_info[0], client_info[1], client_info[2],
                                                               paiement[0], paiement[1], paiement[2], paiement[3]))
                temp_info_list.append((credit_id,) + client_info + paiement)
                total_paiements += 1
                montant_total += paiement[1]

        # Ajouter deux lignes vides
        tree.insert("", "end", values=("", "", "", "", "", "", "", ""))
        tree.insert("", "end", values=("", "", "", "", "", "", "", ""))

        # Ajouter une ligne pour le total des paiements
        tree.insert("", "end", values=("Total", total_paiements, "", "", f" {montant_total} DT", "", "", ""))

        conn.close()

        # Ajouter un espace vide en bas
        espace_vide_bas = Frame(rapport_window, height=20)
        espace_vide_bas.grid(row=2, column=0, columnspan=2, pady=(0, 20), sticky="nsew")

        # Ajouter un espace vide entre le bouton et le bas de la fenêtre
        bottom_frame = tbk.Frame(rapport_window)
        bottom_frame.grid(row=3, column=0, columnspan=2, pady=(20, 0))

        # Bouton d'impression au centre
        tbk.Button(bottom_frame, text="Imprimer",bootstyle="success",
         command=lambda:self.generer_rapport_mensuel_pdf(temp_info_list,mois,str(annee),total_paiements,montant_total)).grid(row=4,column=1,columnspan=1,pady=(20,0))
        print(temp_info_list)
        return (temp_info_list,total_paiements,montant_total)

    def generer_rapport_mensuel_pdf(self,tempo_info_list,mois,annee,total_paiements, montant_total):
        pdfmetrics.registerFont(TTFont('SimplifiedArabic', "Document/NotoNaskhArabic-Regular.ttf"))
        simplified_arabic_style = ParagraphStyle(name='SimplifiedArabic', fontName='SimplifiedArabic',fontSize=14)
        header_style = ParagraphStyle(name='HeaderStyle', fontName='SimplifiedArabic', fontSize=14, alignment=1)
        # Créer un document PDF avec la taille de page A4
        doc = SimpleDocTemplate("rapport_mensuel.pdf", pagesize=A4)

        # Ouvrir et lire le fichier JSON contenant les informations sur l'institution financière
        with open('donnees_imf.json', 'r') as json_file:
            data = json.load(json_file)
            raison_sociale = data['Raison Sociale']

        # Créer une liste pour les données du tableau
        data = [['Id','Nom', 'Prénom', 'CIN', 'Date', 'Montant', 'Reçu', 'Mode']] + tempo_info_list

        # Créer une table pour les données avec une largeur de colonne uniforme
        table = Table(data, colWidths=[50, 80,80,50,80,50,50,80])

        # Appliquer un style à la table
        style = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.gray),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)])

        table.setStyle(style)
        # Créez un style pour le titre
        title_style = ParagraphStyle(name='TitleStyle', fontName='Helvetica-Bold', fontSize=14)

        # Créer le titre du rapport
        title = Paragraph(f'<b>Recouvrement mensuel du mois de {mois}  {annee}</b>', title_style)

        # Créer l'en-tête avec la raison sociale de l'institution financière
        raison_sociale_formated = get_display( arabic_reshaper.reshape(raison_sociale))
        header = Paragraph(f'<b>{raison_sociale_formated}</b>', header_style)

        # Ajoutez un espace vide entre l'en-tête et le titre
        spacer = Spacer(1, 20)
        # Créez les paragraphes pour afficher les valeurs numériques
        total_paragraph = Paragraph(f'<b>Total Paiements: {total_paiements} <br/> <br/> Montant Total: {montant_total} DT </b> ' )


        # Ajoutez un espace vide entre l'en-tête et le titre
        spacer = Spacer(1, 20)

        spacer_2 = Spacer(1, 20)
        saut_ligne = Paragraph(f'<b/><b/><b/><b/><b/> <b/><b/><b/><b/><b/>',getSampleStyleSheet()["Title"])
        # Ajouter le titre et l'en-tête au document
        content = [header,spacer,title,spacer_2,table,spacer_2,total_paragraph]

        if messagebox.askyesno("Impression du rapport", "Voulez-vous imprimer le rapport ?"):
            content.append(spacer)
            # Créer le PDF sans l'imprimer
            pdf_filename = "rapport_mensuel.pdf"
            doc = SimpleDocTemplate(pdf_filename, pagesize=A4)
            doc.build(content)
            # Imprimer le PDF
            self.imprimer_pdf(nom_fichier=pdf_filename)
        else:
            messagebox.showinfo("Impression annulée", "L'impression du rapport a été annulée.")

    def generer_rapport_journalier_pdf(self,tempo_info_list,debut,fin,total_paiements, montant_total):
        pdfmetrics.registerFont(TTFont('SimplifiedArabic', "Document/NotoNaskhArabic-Regular.ttf"))
        simplified_arabic_style = ParagraphStyle(name='SimplifiedArabic', fontName='SimplifiedArabic',fontSize=14)
        header_style = ParagraphStyle(name='HeaderStyle', fontName='SimplifiedArabic', fontSize=14, alignment=1)
        # Créer un document PDF avec la taille de page A4
        doc = SimpleDocTemplate("rapport_journalier.pdf", pagesize=A4)

        # Ouvrir et lire le fichier JSON contenant les informations sur l'institution financière
        with open('donnees_imf.json', 'r') as json_file:
            data = json.load(json_file)
            raison_sociale = data['Raison Sociale']

        # Créer une liste pour les données du tableau
        data = [['Id','Nom', 'Prénom', 'CIN', 'Date', 'Montant', 'Reçu', 'Mode']] + tempo_info_list

        # Créer une table pour les données avec une largeur de colonne uniforme
        table = Table(data, colWidths=[50, 80,80,50,80,50,50,80])

        # Appliquer un style à la table
        style = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.gray),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)])

        table.setStyle(style)
        # Créez un style pour le titre
        title_style = ParagraphStyle(name='TitleStyle', fontName='Helvetica-Bold', fontSize=14)

        # Créer le titre du rapport
        title = Paragraph(f'<b>Recouvrement journalier de {debut} à {fin}</b>', title_style)

        # Créer l'en-tête avec la raison sociale de l'institution financière
        raison_sociale_formated = get_display( arabic_reshaper.reshape(raison_sociale))
        header = Paragraph(f'<b>{raison_sociale_formated}</b>', header_style)

        # Ajoutez un espace vide entre l'en-tête et le titre
        spacer = Spacer(1, 20)
        # Créez les paragraphes pour afficher les valeurs numériques
        total_paragraph = Paragraph(f'<b>Total Paiements: {total_paiements} &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Montant Total: {montant_total} DT </b> ' )


        # Ajoutez un espace vide entre l'en-tête et le titre
        spacer = Spacer(1, 20)

        spacer_2 = Spacer(1, 20)
        saut_ligne = Paragraph(f'<b/><b/><b/><b/><b/> <b/><b/><b/><b/><b/>',getSampleStyleSheet()["Title"])
        # Ajouter le titre et l'en-tête au document
        content = [header,spacer,title,spacer_2,table,spacer_2,total_paragraph]

        # Générer le PDF
        #doc.build(content)
        if messagebox.askyesno("Impression du rapport", "Voulez-vous imprimer le rapport ?"):
            content.append(spacer)
            # Créer le PDF sans l'imprimer
            pdf_filename = "rapport_journalier.pdf"
            doc = SimpleDocTemplate(pdf_filename, pagesize=A4)
            doc.build(content)
            # Imprimer le PDF
            self.imprimer_pdf(nom_fichier=pdf_filename)
        else:
            messagebox.showinfo("Impression annulée", "L'impression du rapport a été annulée.")
    def generer_rapport_global(self):
        # Créer une fenêtre pop-up pour saisir les dates de début et de fin
        popup = Toplevel(self.application)
        popup.title("Générer rapport ")
        popup.geometry("800x600")

        # Labels et champs de saisie pour sélectionner le mois et l'année
        Label(popup, text="Mois :").pack(pady=5)
        self.mois_var = StringVar()
        self.mois_combo = tbk.Combobox(popup, textvariable=self.mois_var, values=self.get_months())
        self.mois_combo.pack()

        Label(popup, text="Année :").pack(pady=5)
        self.annee_var = StringVar()
        self.annee_combo = tbk.Combobox(popup, textvariable=self.annee_var, values=self.get_years())
        self.annee_combo.pack()
        # Ajouter un espace vide entre le bouton et le bas de la fenêtre
        bottom_frame = Frame(popup)
        bottom_frame.pack(side="bottom", pady=20)
        # Bouton pour générer le rapport
        tbk.Button(bottom_frame, text="Générer rapport", bootstyle="success",
                   command=lambda: self.generer_rapport_paye_glob(popup, self.mois_var, self.annee_var)).pack()

    def generer_rapport_paye_glob(self, popup, mois_var, annee_var):
        # Extraire les valeurs sélectionnées des objets StringVar
        mois = mois_var.get().lower()
        annee = int(annee_var.get())

        # Convertir le nom du mois en numéro de mois
        mois_numerique = {mois: idx for idx, mois in enumerate(calendar.month_name) if mois}
        mois_num = mois_numerique.get(mois)

        # Créer le premier jour du mois
        premier_jour = datetime(annee, mois_num, 1).date()

        # Créer le dernier jour du mois
        dernier_jour = (
                    datetime(annee, mois_num % 12 + 1, 1) - timedelta(days=1)).date() if mois_num != 12 else datetime(
            annee, 12, 31).date()

        # Convertir les dates en chaînes de caractères au format "jour mois année"
        dernier_jour_str = dernier_jour.strftime("%Y-%m-%d")
        date_obj = dernier_jour

        print("Premier jour du mois:", premier_jour)
        print("Dernier jour du mois:", dernier_jour_str)

        # Création de la fenêtre principale
        rapport_window = Toplevel(popup)
        rapport_window.title("Rapport Total Paiements")
        rapport_window.geometry("1200x800")

        # Création de la grille
        rapport_window.grid_rowconfigure(0, weight=1)
        rapport_window.grid_rowconfigure(1, weight=1)
        rapport_window.grid_rowconfigure(2, weight=1)
        rapport_window.grid_columnconfigure(0, weight=1)
        rapport_window.grid_columnconfigure(1, weight=1)

        # Titre du rapport centré dans le tableau
        titre_label = Label(rapport_window, text=f"Recouvrement Total jusqu'à {mois} {annee}",
                            font=('Calibri', 14, 'bold'))
        titre_label.grid(row=0, column=1, columnspan=1, pady=(20, 0), sticky="ew")

        style = tbk.Style()
        style.configure("mystyle.Treeview", highlightthickness=1, bd=0, font=('Calibri', 11))
        style.configure("mystyle.Treeview.Heading", font=('Calibri', 12, 'bold'), background="#808080")

        # Création du Treeview
        tree = tbk.Treeview(rapport_window, style="mystyle.Treeview")
        tree["columns"] = ("Nom", "Prénom", "CIN", "Credit", "Total_echeance", "Paiements", "Impayes", "taux")

        # Réduire l'espace entre les colonnes
        tree.column("#0", width=100)
        for col in tree["columns"]:
            tree.column(col, width=100)

        # Aligner les titres des colonnes avec les données
        tree.heading("#0", text="", anchor="w")
        for col in tree["columns"]:
            tree.heading(col, text=col, anchor="w")

        tree.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=20)

        # Récupérer les crédits dans la plage de dates spécifiée
        conn = sqlite3.connect("database.db", detect_types=sqlite3.PARSE_DECLTYPES)
        cursor = conn.cursor()
        cursor.execute("SELECT Date_Credit FROM credits ")
        rows = cursor.fetchall()
        taux=0
        for row in rows:
            if row[0] <= date_obj:
                date_sel = row[0].strftime("%d/%m/%Y")


                cursor.execute(
                    "SELECT Nom, Prenom, CIN, Montant_credit, Total_echeance, Total_paiements, Impayes FROM paiements_total WHERE Date_Credit = ?",
                    (date_sel,))
                results = cursor.fetchall()

                print(results)

                if results:  # Vérifie si la liste de résultats n'est pas vide
                    for ligne in results:
                        print(ligne)
                        taux = f"{round((ligne[5] / ligne[4]), 3) * 100} %"
                        tree.insert("", "end", text="",
                                    values=(ligne[0], ligne[1], ligne[2], ligne[3], ligne[4], ligne[5], ligne[6],taux))

                    # Ajouter un espace vide en bas
                    espace_vide_bas = Frame(rapport_window, height=20)
                    espace_vide_bas.grid(row=2, column=0, columnspan=2, pady=(0, 20), sticky="nsew")
            else:
                pass

        conn.close()

        # Ajouter un espace vide entre le bouton et le bas de la fenêtre
        bottom_frame = tbk.Frame(rapport_window)
        bottom_frame.grid(row=3, column=0, columnspan=2, pady=(20, 0))

        tbk.Button(bottom_frame, text="Imprimer", bootstyle="success").grid(row=4, column=1, columnspan=1,
                                                                            pady=(20, 0))


def generer_rapport_global_pdf(self, tempo_info_list, mois, annee, total_paiements, montant_total):
    pdfmetrics.registerFont(TTFont('SimplifiedArabic', "Document/NotoNaskhArabic-Regular.ttf"))
    simplified_arabic_style = ParagraphStyle(name='SimplifiedArabic', fontName='SimplifiedArabic', fontSize=14)
    header_style = ParagraphStyle(name='HeaderStyle', fontName='SimplifiedArabic', fontSize=14, alignment=1)
    # Créer un document PDF avec la taille de page A4
    doc = SimpleDocTemplate("rapport_mensuel.pdf", pagesize=A4)

    # Ouvrir et lire le fichier JSON contenant les informations sur l'institution financière
    with open('donnees_imf.json', 'r') as json_file:
        data = json.load(json_file)
        raison_sociale = data['Raison Sociale']

    # Créer une liste pour les données du tableau
    data = [['Id', 'Nom', 'Prénom', 'CIN', 'Date', 'Montant', 'Reçu', 'Mode']] + tempo_info_list

    # Créer une table pour les données avec une largeur de colonne uniforme
    table = Table(data, colWidths=[50, 80, 80, 50, 80, 50, 50, 80])

    # Appliquer un style à la table
    style = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.gray),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black)])

    table.setStyle(style)
    # Créez un style pour le titre
    title_style = ParagraphStyle(name='TitleStyle', fontName='Helvetica-Bold', fontSize=14)

    # Créer le titre du rapport
    title = Paragraph(f'<b>Recouvrement mensuel du mois de {mois}  {annee}</b>', title_style)

    # Créer l'en-tête avec la raison sociale de l'institution financière
    raison_sociale_formated = get_display(arabic_reshaper.reshape(raison_sociale))
    header = Paragraph(f'<b>{raison_sociale_formated}</b>', header_style)

    # Ajoutez un espace vide entre l'en-tête et le titre
    spacer = Spacer(1, 20)
    # Créez les paragraphes pour afficher les valeurs numériques
    total_paragraph = Paragraph(
        f'<b>Total Paiements: {total_paiements} <br/> <br/> Montant Total: {montant_total} DT </b> ')

    # Ajoutez un espace vide entre l'en-tête et le titre
    spacer = Spacer(1, 20)

    spacer_2 = Spacer(1, 20)
    saut_ligne = Paragraph(f'<b/><b/><b/><b/><b/> <b/><b/><b/><b/><b/>', getSampleStyleSheet()["Title"])
    # Ajouter le titre et l'en-tête au document
    content = [header, spacer, title, spacer_2, table, spacer_2, total_paragraph]

    if messagebox.askyesno("Impression du rapport", "Voulez-vous imprimer le rapport ?"):
        content.append(spacer)
        # Créer le PDF sans l'imprimer
        pdf_filename = "rapport_mensuel.pdf"
        doc = SimpleDocTemplate(pdf_filename, pagesize=A4)
        doc.build(content)
        # Imprimer le PDF
        self.imprimer_pdf(nom_fichier=pdf_filename)
    else:
        messagebox.showinfo("Impression annulée", "L'impression du rapport a été annulée.")

if __name__ == "__main__":
    app = Application()
    app.mainloop()

