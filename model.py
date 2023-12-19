import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
from PIL import Image, ImageTk
import time
import ctypes

ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

class ProgramGUI:
    pn_condition_1 = "10449850"
    pn_condition_2 = "10449565"
    pn_condition_3 = "10478845"
    pn_condition_4 = "10403646"
    pn_condition_5 = "10286094"
    pn_condition_6 = "10408736"
    pn_condition_7 = "10408735"
    pn_condition_8 = "10408737"
    pn_condition_9 = "20003697"
    pn_condition_10 = "10435298"
    pn_condition_11 = "10484970"
    pn_condition_12 = "10503913"
    pn_condition_13 = "10503915"
    pn_condition_14 = "10466155"
    pn_condition_15 = "10466153"
    pn_condition_16 = "10422741"
    pn_condition_17 = "10511135"
    pn_condition_18 = "10460520"
    pn_condition_19 = "10488446"
    pn_condition_20 = "10476214"
    pn_condition_21 = "10467795"
    pn_condition_22 = "10476227"
    pn_condition_23 = "10488988"
    pn_condition_24 = "10511006"
    pn_condition_25 = "10511005"
    pn_condition_26 = "10511134"
    pn_condition_27 = "10404220"
    pn_condition_28 = "10494427"
    pn_condition_29 = "10489031"
    pn_condition_30 = "10449851"
    pn_condition_31 = "10511323"
    pn_condition_32 = "10451100"
    pn_condition_33 = "10451241"
    pn_condition_34 = "10466769"
    pn_condition_35 = "10449848"
    pn_condition_36 = "10398444"
    pn_condition_37 = "10405144"
    pn_condition_38 = "10363040"
    pn_condition_39 = "10476648"
    pn_condition_40 = "10476304"
    pn_condition_41 = "10418254"
    pn_condition_42 = "10450760"
    pn_condition_43 = "10476124"
    pn_condition_44 = "10365875"
    pn_condition_45 = "10511151"
    pn_condition_46 = "10363039"
    pn_condition_47 = "10471263"
    pn_condition_48 = "10480112"
    pn_condition_49 = "10462063"
    pn_condition_50 = "10484681"
    pn_condition_51 = "10466777"
    pn_condition_52 = "10478700"
    pn_condition_53 = "10327190"
    pn_condition_54 = "10494306"
    pn_condition_55 = "10450762"
    pn_condition_56 = "10480116"
    pn_condition_57 = "10480114"
    pn_condition_58 = "10327186"
    pn_condition_59 = "10459295"
    pn_condition_60 = "10365879"
    pn_condition_61 = "10451334"
    pn_condition_62 = "10465177"
    pn_condition_63 = "10506371"
    pn_condition_64 = "10506370"
    pn_condition_65 = "10478700"
    pn_condition_66 = "10441451"
    pn_condition_67 = "10327204"
    pn_condition_68 = "10365880"
    pn_condition_69 = "10384290"
    pn_condition_70 = "10428460"
    pn_condition_71 = "10441452"
    pn_condition_72 = "10496775"
    pn_condition_73 = "10465178"
    pn_condition_74 = "10489545"
    pn_condition_75 = "10432252"
    pn_condition_76 = "10418252"
    pn_condition_77 = "10324235"
    pn_condition_78 = "10455632"
    pn_condition_79 = "10417838"
    pn_condition_80 = "10442174"
    pn_condition_81 = "10369596"
    pn_condition_82 = "10419762"
    pn_condition_83 = "10453899"
    pn_condition_84 = "10360819"
    pn_condition_85 = "10468161"
    pn_condition_86 = "10511322"
    pn_condition_87 = "10511324"
    pn_condition_88 = "10324236"
    pn_condition_89 = "10511325"
    pn_condition_90 = "10476460"
    pn_condition_91 = "10442176"
    pn_condition_92 = "10508247"
    pn_condition_93 = "10319139"
    pn_condition_94 = "10469505"
    pn_condition_95 = "10466768"
    pn_condition_96 = "10384283"
    pn_condition_97 = "10449564"
    pn_condition_98 = "10472088"
    pn_condition_99 = "10465174"
    pn_condition_100 = "10384289"
    pn_condition_101 = "10483789"
    pn_condition_102 = "10471140"
    pn_condition_103 = "10454702"
    pn_condition_104 = "10480134"
    pn_condition_105 = "10384281"
    pn_condition_106 = "10514000"
    pn_condition_107 = "10327187"
    pn_condition_108 = "10432893"
    pn_condition_109 = "10492959"
    pn_condition_110 = "10419761"
    pn_condition_111 = "10393645"
    pn_condition_112 = "10449559"
    pn_condition_113 = "10418255"
    pn_condition_114 = "10450759"
    pn_condition_115 = "20003703"
    pn_condition_116 = "10480110"
    pn_condition_117 = "10426934"
    pn_condition_118 = "10467943"
    pn_condition_119 = "10488205"
    pn_condition_120 = "10483868"
    pn_condition_121 = "10433007"
    pn_condition_122 = "10480124"
    pn_condition_123 = "10471550"
    pn_condition_124 = "10450990"
    pn_condition_125 = "10380072"
    pn_condition_126 = "10359986"
    pn_condition_127 = "10380613"
    pn_condition_128 = "20003708"
    pn_condition_129 = "10451608"
    pn_condition_130 = "10489692"
    pn_condition_131 = "10480108"
    pn_condition_132 = "10509018"
    pn_condition_133 = "10429500"
    pn_condition_134 = "10463307"
    pn_condition_135 = "10367857"
    pn_condition_136 = "20003706"
    pn_condition_137 = "10466155"
    pn_condition_138 = "10520705"
    pn_condition_139 = "10520704"
    pn_condition_140 = "10496773"
    pn_condition_141 = "10451609"
    pn_condition_142 = "20003697"
    pn_condition_143 = "10508938"
    pn_condition_144 = "10467683"
    pn_condition_145 = "10319133"
    pn_condition_146 = "10462619"
    pn_condition_147 = "10489325"
    pn_condition_148 = "10466153"
    pn_condition_149 = "10442172"
    pn_condition_150 = "10384243"
    pn_condition_151 = "20003703"
    pn_condition_152 = "10442174"
    pn_condition_153 = "20002946"
    pn_condition_154 = "10513617"
    pn_condition_155 = "10508500"
    pn_condition_156 = "10442176"
    pn_condition_157 = "10492722"
    pn_condition_158 = "10398463"
    pn_condition_159 = "10475604"
    pn_condition_160 = "10396908"
    pn_condition_161 = "10508258"
    pn_condition_162 = "10508175"
    pn_condition_163 = "10480120"
    pn_condition_164 = "10476136"
    pn_condition_165 = "10478863"
    pn_condition_166 = "10463247"
    pn_condition_167 = "10429497"
    pn_condition_168 = "10334955"
    pn_condition_169 = "10498223"
    pn_condition_170 = "10393780"
    pn_condition_171 = "10511323"
    pn_condition_172 = "10508226"
    pn_condition_173 = "10489706"
    pn_condition_174 = "10476130"
    pn_condition_175 = "10480118"
    pn_condition_176 = "10428442"
    pn_condition_177 = "10431602"
    pn_condition_178 = "10431603"
    pn_condition_179 = "10492982"
    pn_condition_180 = "10502865"
    pn_condition_181 = "10484789"
    pn_condition_182 = "10502864"
    pn_condition_183 = "10484790"
    pn_condition_184 = "10511507"
    pn_condition_185 = "10476054"
    pn_condition_186 = "10524054"
    pn_condition_187 = "10511498"
    pn_condition_188 = "10511508"
    pn_condition_189 = "10524051"
    pn_condition_190 = "10501813"
    pn_condition_191 = "10467796"
    pn_condition_192 = "10428444"
    pn_condition_193 = "10524052"
    pn_condition_194 = "10418253"
    pn_condition_195 = "10319133"
    pn_condition_196 = "10524049"
    pn_condition_197 = "10519223"
    pn_condition_198 = "10499064"
    pn_condition_199 = "10519223"
    pn_condition_200 = "10519223"
    pn_condition_201 = "10519235"
    pn_condition_202 = "10511660"
    pn_condition_203 = "10511661"
    pn_condition_204 = "10319137"
    pn_condition_205 = "10508941"
    pn_condition_206 = "10509678"
    pn_condition_207 = "10484790"
    pn_condition_208 = "10319138"
    pn_condition_209 = "20002946"
    pn_condition_210 = "10286095"
    pn_condition_211 = "10521757"
    pn_condition_212 = "10508947"
    pn_condition_213 = "10530568"
    pn_condition_214 = "10521046"
    pn_condition_215 = "10513786"
    pn_condition_216 = "10521799"
    pn_condition_217 = "10521047"
    pn_condition_218 = "10526383"
    pn_condition_219 = "10517617"
    pn_condition_220 = "10517614"
    pn_condition_221 = "10451333"
    pn_condition_222 = "10499221"
    pn_condition_223 = "10517614"
    pn_condition_224 = "10517617"
    pn_condition_225 = "10428640"
    pn_condition_226 = "10466771"
    pn_condition_227 = "10465175"
    pn_condition_228 = "10494291"
    pn_condition_229 = "10536087"
    pn_condition_230 = "10454702"
    pn_condition_231 = "10524053"
    pn_condition_232 = "10506373"
    pn_condition_233 = "10476052"
    pn_condition_234 = "10462062"
    pn_condition_235 = "10475852"
    pn_condition_236 = "10515012"
    pn_condition_237 = "10506372"
    pn_condition_238 = "10511590"
    pn_condition_239 = "10511591"
    pn_condition_240 = "10526027"
    pn_condition_241 = "10511355"
    pn_condition_242 = "10511724"
    pn_condition_243 = "10511723"
    pn_condition_244 = "10519851"
    pn_condition_245 = "10509846"
    pn_condition_246 = "10520892"
    pn_condition_247 = "10517647"
    pn_condition_248 = "10502066"
    pn_condition_249 = "10509878"
    pn_condition_250 = "10519112"
    pn_condition_251 = "10511329"
    pn_condition_252 = "10302232"
    pn_condition_253 = "10302233"
    pn_condition_254 = "10524050"
    pn_condition_255 = "10511588"
    pn_condition_256 = "10511589"
    pn_condition_257 = "10505549"
    pn_condition_258 = "10505551"
    pn_condition_259 = "10530524"
    pn_condition_260 = "10398443"
    pn_condition_261 = "10442175"
    pn_condition_262 = "10465176"
    pn_condition_263 = "10526177"
    pn_condition_264 = "10503903"
    pn_condition_265 = "10519233"
    pn_condition_266 = "10466758"
    pn_condition_267 = "10519373"
    pn_condition_268 = "10519416"
    pn_condition_269 = "10515012"
    pn_condition_270 = "10532260"
    pn_condition_271 = "10515013"
    pn_condition_272 = "10524376"
    pn_condition_273 = "10320428"
    pn_condition_274 = "10320486"
    pn_condition_275 = "10476140"
    pn_condition_276 = "10415305"
    pn_condition_277 = "10414540"
    pn_condition_278 = "10444910"
    pn_condition_279 = "10449839"
    pn_condition_280 = "20002949"
    pn_condition_281 = "10486269"
    pn_condition_282 = "10533713"

    def __init__(self, master):
        self.master = master
        self.master.title("Spis sztuk NOK - SL")
        self.master.geometry("900x750")
        self.master.configure(bg='navy')
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

        img_path = os.path.join(os.path.dirname(__file__), "gkn.png")
        img = Image.open(img_path)
        img = img.resize((200, 50))
        img = ImageTk.PhotoImage(img)
        self.label_image = tk.Label(master, image=img, bg='navy')
        self.label_image.image = img
        self.label_image.pack(anchor='nw', padx=10, pady=10)

        center_alignment = "nsew"

        self.label_pn = tk.Label(master, text="Numer Części:", font=("Arial", 12, "bold"), bg='navy', fg='gold')
        self.label_pn.pack(pady=(10, 0))

        self.entry_pn = tk.Entry(master, font=("Arial", 12))
        self.entry_pn.pack()

        self.label_container = tk.Label(master, text="Numer pojemnika:", font=("Arial", 12, "bold"), bg='navy', fg='gold')
        self.label_container.pack(pady=(10, 0))

        self.entry_container = tk.Entry(master, font=("Arial", 12))
        self.entry_container.pack()
        self.entry_container.bind('<FocusOut>', self.validate_numer_pojemnika)

        self.label_quantity = tk.Label(master, text="Ilość sztuk NOK:", font=("Arial", 12, "bold"), bg='navy', fg='gold')
        self.label_quantity.pack(pady=(10, 0))

        vcmd = (master.register(self.validate_ilosc_sztuk), '%P')
        self.entry_quantity = tk.Entry(master, validate="key", validatecommand=vcmd, font=("Arial", 12))
        self.entry_quantity.pack()
        self.entry_quantity.bind('<FocusOut>', self.update_table)

        self.label_powod_nok = tk.Label(master, text="Powód sztuk NOK:", font=("Arial", 12, "bold"), bg='navy', fg='gold')
        self.label_powod_nok.pack(pady=(10, 0))

        self.entry_powod_nok = tk.Entry(master, font=("Arial", 12))
        self.entry_powod_nok.pack()

        self.label_dzialania = tk.Label(master, text="Działania:", font=("Arial", 12, "bold"), bg='navy', fg='gold')
        self.label_dzialania.pack(pady=(10, 0))

        self.slider_dzialania = ttk.Combobox(master, values=["Przeróbka", "Złom", "Dopuszczenie", "Reklamacja"], state="readonly")
        self.slider_dzialania.pack()

        self.label_data = tk.Label(master, text="Data:", font=("Arial", 12, "bold"), bg='navy', fg='gold')
        self.label_data.pack(pady=(10, 0))

        self.label_current_date = tk.Label(master, text="", font=("Arial", 12), bg='navy', fg='white')
        self.label_current_date.pack()

        self.label_godzina = tk.Label(master, text="Godzina:", font=("Arial", 12, "bold"), bg='navy', fg='gold')
        self.label_godzina.pack(pady=(10, 0))

        self.label_current_time = tk.Label(master, text="", font=("Arial", 12), bg='navy', fg='white')
        self.label_current_time.pack()

        self.table_frame = tk.Frame(master)
        self.table_frame.pack(side="top")

        self.table_entries = []
        self.original_table_values = []
        self.table_entry_active = False  # Nowa linijka

        self.update_datetime()

        self.button_verify = tk.Button(master, text="Weryfikuj", command=self.weryfikuj, font=("Arial", 12, "bold"), fg='black', bg='skyblue')
        self.button_verify.pack(pady=10)

        self.button_add = tk.Button(master, text="Dodaj dane", command=self.dodaj_dane, font=("Arial", 12, "bold"), fg='black', bg='skyblue', state="disabled")
        self.button_add.pack(pady=10)

        self.button_legend = tk.Button(master, text="Legenda", command=self.show_legend, font=("Arial", 12, "bold"), fg='black', bg='skyblue')
        self.button_legend.pack(side="bottom", anchor="se", padx=10, pady=10)


    def show_legend(self):
        # Stworzenie nowego okna dla legendy
        legend_window = tk.Toplevel(self.master)
        legend_window.title("Legenda")
        legend_window.geometry("300x180")

        # Wczytanie obrazu legendy
        img_path = os.path.join(os.path.dirname(__file__), "legenda.jpg")
        img = Image.open(img_path)
        img = ImageTk.PhotoImage(img)

        # Dodanie etykiety z obrazem do nowego okna
        label_legend = tk.Label(legend_window, image=img, bg='white')
        label_legend.image = img
        label_legend.pack(expand=True, fill="both")

    def on_table_entry_focus_in(self, entry):
        self.table_entry_active = True
        entry.configure(bg='yellow')
        self.check_yellow_entries()

    def check_yellow_entries(self):
        if any(entry.cget("background") == 'yellow' for entry in self.table_entries):
            self.button_add["state"] = "disabled"

    def on_table_entry_focus_out(self, entry):
        self.table_entry_active = False
        self.on_table_entry_change(entry)

    def handle_last_entry_enter(self, event):
        self.button_add.focus()

    def validate_ilosc_sztuk(self, new_value):
        return not new_value or new_value.isdigit()

    def update_table(self, event):
        ilosc_sztuk_value = self.entry_quantity.get()

        if not ilosc_sztuk_value or ilosc_sztuk_value.isdigit():
            ilosc_sztuk = int(ilosc_sztuk_value) if ilosc_sztuk_value else 0
        else:
            messagebox.showwarning("Ostrzeżenie", "Wprowadzona ilość sztuk nie jest liczbą!")
            self.entry_quantity.delete(0, tk.END)
            return

        if ilosc_sztuk > 100:
            messagebox.showwarning("Ostrzeżenie", "Zadeklarowano za dużą ilość sztuk!")
            self.entry_quantity.delete(0, tk.END)
            return

        for entry in self.table_entries:
            entry.destroy()

        self.table_entries = []
        self.original_table_values = []

        if ilosc_sztuk > 0:
            for i in range(ilosc_sztuk):
                entry = tk.Entry(self.table_frame, width=5, font=("Arial", 12))
                entry.grid(row=i // 10, column=i % 10, padx=5, pady=5)
                entry.bind('<Return>', lambda e, row=i // 10, col=i % 10, self=self: self.move_to_next_entry(row, col))
                entry.bind('<FocusOut>', lambda e, entry=entry, self=self: self.on_table_entry_focus_out(entry))
                entry.bind('<FocusIn>', lambda e, entry=entry, self=self: self.on_table_entry_focus_in(entry))
                self.table_entries.append(entry)
                self.original_table_values.append('')

                # Dodaj kod do obsługi Enter dla ostatniego pola tabeli
                if i == ilosc_sztuk - 1:
                    entry.bind('<Return>', lambda e, self=self: self.handle_last_entry_enter(e))

            self.table_frame.pack()
        else:
            self.table_frame.pack_forget()

    def move_to_next_entry(self, row, col):
        next_col = (col + 1) % 10
        next_row = row + 1 if next_col == 0 else row
        self.table_entries[next_row * 10 + next_col].focus()

    def on_table_entry_change(self, entry):
        new_value = entry.get()
        original_value = self.original_table_values[self.table_entries.index(entry)]

        if new_value != original_value:
            if new_value not in self.original_table_values:
                entry.configure(bg='aquamarine')  # Nowa, unikalna wartość
            else:
                entry.configure(bg='orange')  # Wartość powtórzona
            self.button_add["state"] = "disabled"
        else:
            entry.configure(bg='grey')  # Wartość bez zmian
            self.button_add["state"] = "disabled"
        self.check_all_entries_green()

    def check_all_entries_green(self):
        if all(entry.cget("background") == 'green' for entry in self.table_entries):
            self.button_add["state"] = "normal"

    def weryfikuj(self):
        pn_value = self.entry_pn.get()

        if len(pn_value) not in (11, 14):
            messagebox.showwarning("Ostrzeżenie", "Nieprawidłowa ilość znaków w PN. Wprowadź 11 lub 14 znaków.")
            self.entry_pn.delete(0, tk.END)
            return

        if self.pn_condition_1 in pn_value:
            self.table_condition = "124735033042"
        elif self.pn_condition_2 in pn_value:
            self.table_condition = "11773501902"
        elif self.pn_condition_3 in pn_value:
            self.table_condition = "10478845#A#"
        elif self.pn_condition_4 in pn_value:
            self.table_condition = "16GKB2F20"
        elif self.pn_condition_5 in pn_value:
            self.table_condition = "GX73-4K139-DC"
        elif self.pn_condition_6 in pn_value:
            self.table_condition = "M8E2-4K139-AB"
        elif self.pn_condition_7 in pn_value:
            self.table_condition = "M8E2-4K138-AB"
        elif self.pn_condition_8 in pn_value:
            self.table_condition = "M8E2-4K139-BB"
        elif self.pn_condition_9 in pn_value:
            self.table_condition = "HK83-3N129-AC"
        elif self.pn_condition_10 in pn_value:
            self.table_condition = "AZ5A1486202"
        elif self.pn_condition_11 in pn_value:
            self.table_condition = "#95B407271K"
        elif self.pn_condition_12 in pn_value:
            self.table_condition = "16GKU9C20"
        elif self.pn_condition_13 in pn_value:
           self.table_condition = "16GKU9F10"
        elif self.pn_condition_14 in pn_value:
            self.table_condition = "L8D2-4K138-AF"
        elif self.pn_condition_15 in pn_value:
            self.table_condition = "L8D2-4K139-AE"
        elif self.pn_condition_16 in pn_value:
            self.table_condition = "AY5A1486102"
        elif self.pn_condition_17 in pn_value:
            self.table_condition = "AZ5A733B4"
        elif self.pn_condition_18 in pn_value:
            self.table_condition = "10460520"
        elif self.pn_condition_19 in pn_value:
            self.table_condition = "10488446#"
        elif self.pn_condition_20 in pn_value:
            self.table_condition = "[)>06P32339220T"
        elif self.pn_condition_21 in pn_value:
            self.table_condition = "#7T0407271A"
        elif self.pn_condition_22 in pn_value:
            self.table_condition = "[)>06P32339218T"
        elif self.pn_condition_23 in pn_value:
            self.table_condition = "16GKU2E90"
        elif self.pn_condition_24 in pn_value:
            self.table_condition = "AZ5A4C69201"
        elif self.pn_condition_25 in pn_value:
            self.table_condition = "AY5A4C69101"
        elif self.pn_condition_26 in pn_value:
            self.table_condition = "AY5A733B3"
        elif self.pn_condition_27 in pn_value:
            self.table_condition = "AZ9896324"
        elif self.pn_condition_28 in pn_value:
            self.table_condition = "10494427"
        elif self.pn_condition_29 in pn_value:
            self.table_condition = "16GKU2E80"
        elif self.pn_condition_30 in pn_value:
            self.table_condition = "12473503404"
        elif self.pn_condition_31 in pn_value:
            self.table_condition = "#95B501204A"
        elif self.pn_condition_32 in pn_value:
            self.table_condition = "12933300600"
        elif self.pn_condition_33 in pn_value:
            table_condition = "505639970028203"
        elif self.pn_condition_34 in pn_value:
            self.table_condition = "10466769#"
        elif self.pn_condition_35 in pn_value:
            self.table_condition = "11773502102"
        elif self.pn_condition_36 in pn_value:
            self.table_condition = "11773302801"
        elif self.pn_condition_37 in pn_value:
            self.table_condition = "AZ808985602"
        elif self.pn_condition_38 in pn_value:
            self.table_condition = "AZ785693604"
        elif self.pn_condition_39 in pn_value:
            self.table_condition = "#5WA501203A"
        elif self.pn_condition_40 in pn_value:
            self.table_condition = "P1188165-00-A"
        elif self.pn_condition_41 in pn_value:
            self.table_condition = "11183503300"
        elif self.pn_condition_42 in pn_value:
            self.table_condition = "11773305902"
        elif self.pn_condition_43 in pn_value:
            self.table_condition = "#4M0501201D"
        elif self.pn_condition_44 in pn_value:
            self.table_condition = "AY8666073"
        elif self.pn_condition_45 in pn_value:
            self.table_condition = "AZ5A733B6"
        elif self.pn_condition_46 in pn_value:
            self.table_condition = "AY785693704"
        elif self.pn_condition_47 in pn_value:
            self.table_condition = "#4K0407271G"
        elif self.pn_condition_48 in pn_value:
            self.table_condition = "#8W0407271Q"
        elif self.pn_condition_49 in pn_value:
            self.table_condition = "10462063"
        elif self.pn_condition_50 in pn_value:
            self.table_condition = "#95B407271L"
        elif self.pn_condition_51 in pn_value:
            self.table_condition = "10466777#A#"
        elif self.pn_condition_52 in pn_value:
            self.table_condition = "[)>0632339110T"
        elif self.pn_condition_53 in pn_value:
            self.table_condition = "AZ866404806"
        elif self.pn_condition_54 in pn_value:
            self.table_condition = "10494306"
        elif self.pn_condition_55 in pn_value:
            self.table_condition = "11773305702"
        elif self.pn_condition_56 in pn_value:
            self.table_condition = "#4M0407271AE"
        elif self.pn_condition_57 in pn_value:
            self.table_condition = "#4M0407271AD"
        elif self.pn_condition_58 in pn_value:
            self.table_condition = "AZ866462405"
        elif self.pn_condition_59 in pn_value:
            self.table_condition = "AY808985503"
        elif self.pn_condition_60 in pn_value:
            self.table_condition = "AY8681973"
        elif self.pn_condition_61 in pn_value:
            self.table_condition = "505602760028203"
        elif self.pn_condition_62 in pn_value:
            self.table_condition = "#4M0407271AA"
        elif self.pn_condition_63 in pn_value:
            self.table_condition = "#2N0407271R"
        elif self.pn_condition_64 in pn_value:
            self.table_condition = "#2N0407271Q"
        elif self.pn_condition_65 in pn_value:
            self.table_condition = "[)>06P32339110T"
        elif self.pn_condition_66 in pn_value:
            self.table_condition = "11773303202"
        elif self.pn_condition_67 in pn_value:
            self.table_condition = "AZ866404606"
        elif self.pn_condition_68 in pn_value:
            self.table_condition = "AZ8681972"
        elif self.pn_condition_69 in pn_value:
            self.table_condition = "AZ8689556"
        elif self.pn_condition_70 in pn_value:
            self.table_condition = "16GKU8740"
        elif self.pn_condition_71 in pn_value:
            self.table_condition = "11773303602"
        elif self.pn_condition_72 in pn_value:
            self.table_condition = "#1N3501203B"
        elif self.pn_condition_73 in pn_value:
            self.table_condition = "#4M0407271AB"
        elif self.pn_condition_74 in pn_value:
            self.table_condition = "#8B3407271B"
        elif self.pn_condition_75 in pn_value:
            self.table_condition = "11773300802"
        elif self.pn_condition_76 in pn_value:
            self.table_condition = "11773505301"
        elif self.pn_condition_77 in pn_value:
            self.table_condition = "AY863945906"
        elif self.pn_condition_78 in pn_value:
            self.table_condition = "AZ5A0F09802"
        elif self.pn_condition_79 in pn_value:
            self.table_condition = "16GKU9060"
        elif self.pn_condition_80 in pn_value:
            self.table_condition = "K8D2-3B436-CC"
        elif self.pn_condition_81 in pn_value:
            self.table_condition = "12463301801"
        elif self.pn_condition_82 in pn_value:
            self.table_condition = "AZ9423550"
        elif self.pn_condition_83 in pn_value:
            self.table_condition = "6701569260028200"
        elif self.pn_condition_84 in pn_value:
            self.table_condition = "11773305300"
        elif self.pn_condition_85 in pn_value:
            self.table_condition = "#7T0407272A"
        elif self.pn_condition_86 in pn_value:
            self.table_condition = "#95B501204     ###"
        elif self.pn_condition_87 in pn_value:
            self.table_condition = "#95B501204B"
        elif self.pn_condition_88 in pn_value:
            self.table_condition = "AZ863946006"
        elif self.pn_condition_89 in pn_value:
            self.table_condition = "#95B501204C"
        elif self.pn_condition_90 in pn_value:
            self.table_condition = "[)>06P32339223T"
        elif self.pn_condition_91 in pn_value:
            self.table_condition = "L8D2-3B436-AD"
        elif self.pn_condition_92 in pn_value:
            self.table_condition = "#95C501201M"
        elif self.pn_condition_93 in pn_value:
            self.table_condition = "DW93-4K139-AD"
        elif self.pn_condition_94 in pn_value:
            self.table_condition = "10469505#B#"
        elif self.pn_condition_95 in pn_value:
            self.table_condition = "10466768#A#"
        elif self.pn_condition_96 in pn_value:
            self.table_condition = "AY8689553"
        elif self.pn_condition_97 in pn_value:
            self.table_condition = "11773501702"
        elif self.pn_condition_98 in pn_value:
            self.table_condition = "10472088"
        elif self.pn_condition_99 in pn_value:
            self.table_condition = "#4M0407271R"
        elif self.pn_condition_100 in pn_value:
            self.table_condition = "AY8689555"
        elif self.pn_condition_101 in pn_value:
            self.table_condition = "6702173470"
        elif self.pn_condition_102 in pn_value:
            self.table_condition = "#4N0407271J"
        elif self.pn_condition_103 in pn_value:
            self.table_condition = "670158933002820"
        elif self.pn_condition_104 in pn_value:
            self.table_condition = "#4M4407271G"
        elif self.pn_condition_105 in pn_value:
            self.table_condition = "AY8689549"
        elif self.pn_condition_106 in pn_value:
            self.table_condition = "#7T0501201"
        elif self.pn_condition_107 in pn_value:
            self.table_condition = "AY866462505"
        elif self.pn_condition_108 in pn_value:
            self.table_condition = "11773300902"
        elif self.pn_condition_109 in pn_value:
            self.table_condition = "11773303103"
        elif self.pn_condition_110 in pn_value:
            self.table_condition = "AY9423549"
        elif self.pn_condition_111 in pn_value:
            self.table_condition = "11773309700"
        elif self.pn_condition_112 in pn_value:
            self.table_condition = "11773501802"
        elif self.pn_condition_113 in pn_value:
            self.table_condition = "11183503400"
        elif self.pn_condition_114 in pn_value:
            self.table_condition = "11773305802"
        elif self.pn_condition_115 in pn_value:
            self.table_condition = "HK83-3N128-AC"
        elif self.pn_condition_116 in pn_value:
            self.table_condition = "#8W0407271P"
        elif self.pn_condition_117 in pn_value:
            self.table_condition = "AZ9452904"
        elif self.pn_condition_118 in pn_value:
            self.table_condition = "#7T0407272E"
        elif self.pn_condition_119 in pn_value:
            self.table_condition = "10488205"
        elif self.pn_condition_120 in pn_value:
            self.table_condition = "6702173480"
        elif self.pn_condition_121 in pn_value:
            self.table_condition = "11773301002"
        elif self.pn_condition_122 in pn_value:
            self.table_condition = "#4K0407271P"
        elif self.pn_condition_123 in pn_value:
            self.table_condition = "#2N0407272Q"
        elif self.pn_condition_124 in pn_value:
            self.table_condition = "M8E2-3N128-AC23"
        elif self.pn_condition_125 in pn_value:
           self.table_condition = "11773302000"
        elif self.pn_condition_126 in pn_value:
            self.table_condition = "11773300700"
        elif self.pn_condition_127 in pn_value:
            self.table_condition = "505504250028203"
        elif self.pn_condition_128 in pn_value:
            self.table_condition = "HK83-4K138-AB"
        elif self.pn_condition_129 in pn_value:
            self.table_condition = "12323509400"
        elif self.pn_condition_130 in pn_value:
            self.table_condition = "#971501201AA"
        elif self.pn_condition_131 in pn_value:
            self.table_condition = "#8W0407271M"
        elif self.pn_condition_132 in pn_value:
            self.table_condition = "11773303503"
        elif self.pn_condition_133 in pn_value:
            self.table_condition = "16GKU8060"
        elif self.pn_condition_134 in pn_value:
            self.table_condition = "#2N0407272R"
        elif self.pn_condition_135 in pn_value:
            self.table_condition = "#7E0407271AC"
        elif self.pn_condition_136 in pn_value:
            self.table_condition = "HK83-4K139-AB"
        elif self.pn_condition_137 in pn_value:
            self.table_condition = "L8D2-4K138-AF23"
        elif self.pn_condition_138 in pn_value:
            self.table_condition = "11773304303"
        elif self.pn_condition_139 in pn_value:
            self.table_condition = "11773304203"
        elif self.pn_condition_140 in pn_value:
            self.table_condition = "#1N3501204B"
        elif self.pn_condition_141 in pn_value:
            self.table_condition = "12323509300"
        elif self.pn_condition_142 in pn_value:
            self.table_condition = "HK83-3N129-AC23"
        elif self.pn_condition_143 in pn_value:
            self.table_condition = "11773303803"
        elif self.pn_condition_144 in pn_value:
            self.table_condition = "#7T0407271"
        elif self.pn_condition_145 in pn_value:
            self.table_condition = "DW93-3N129-AC"
        elif self.pn_condition_146 in pn_value:
            self.table_condition = "6701589530"
        elif self.pn_condition_147 in pn_value:
            self.table_condition = "#8B3407271A"
        elif self.pn_condition_148 in pn_value:
            self.table_condition = "L8D2-4K139-AE23"
        elif self.pn_condition_149 in pn_value:
            self.table_condition = "EJ32-3B437-CE"
        elif self.pn_condition_150 in pn_value:
            self.table_condition = "AY809219501"
        elif self.pn_condition_151 in pn_value:
            self.table_condition = "HK83-3N128-AC23"
        elif self.pn_condition_152 in pn_value:
            self.table_condition = "K8D2-3B436-CC23"
        elif self.pn_condition_153 in pn_value:
            self.table_condition = "GX53-3N129-AC"
        elif self.pn_condition_154 in pn_value:
            self.table_condition = "46355859"
        elif self.pn_condition_155 in pn_value:
            self.table_condition = "46355858"
        elif self.pn_condition_156 in pn_value:
            self.table_condition = "L8D2-3B436-AD23"
        elif self.pn_condition_157 in pn_value:
            self.table_condition = "11773300303"
        elif self.pn_condition_158 in pn_value:
            self.table_condition = "11773303701"
        elif self.pn_condition_159 in pn_value:
           self.table_condition = "11773309502"
        elif self.pn_condition_160 in pn_value:
            self.table_condition = "16GKU8050"
        elif self.pn_condition_161 in pn_value:
            self.table_condition = "#95C501201P"
        elif self.pn_condition_162 in pn_value:
            self.table_condition = "#95C501201N"
        elif self.pn_condition_163 in pn_value:
            self.table_condition = "#4N0407271N"
        elif self.pn_condition_164 in pn_value:
            self.table_condition = "#4M0501201F"
        elif self.pn_condition_165 in pn_value:
            self.table_condition = "[)>06P32339111T"
        elif self.pn_condition_166 in pn_value:
            self.table_condition = "12233304303"
        elif self.pn_condition_167 in pn_value:
            self.table_condition = "16GKU8060"
        elif self.pn_condition_168 in pn_value:
            self.table_condition = "12533307900"
        elif self.pn_condition_169 in pn_value:
            self.table_condition = "P1188115-99-A"
        elif self.pn_condition_170 in pn_value:
            self.table_condition = "11773309900"
        elif self.pn_condition_171 in pn_value:
            self.table_condition = "#95B501204A    ###"
        elif self.pn_condition_172 in pn_value:
            self.table_condition = "#95C501201Q"
        elif self.pn_condition_173 in pn_value:
            self.table_condition = "#95B501203Q"
        elif self.pn_condition_174 in pn_value:
            self.table_condition = "#4M4501203N"
        elif self.pn_condition_175 in pn_value:
            self.table_condition = "#4M0407271AC"
        elif self.pn_condition_176 in pn_value:
            self.table_condition = "14633300702"
        elif self.pn_condition_177 in pn_value:
            self.table_condition = "521219890"
        elif self.pn_condition_178 in pn_value:
            self.table_condition = "521219880"
        elif self.pn_condition_179 in pn_value:
            self.table_condition = "11773300103"
        elif self.pn_condition_180 in pn_value:
            self.table_condition = "S8E2-4K138-AAVB"
        elif self.pn_condition_181 in pn_value:
            self.table_condition = "S8E2-3N129-AAVB"
        elif self.pn_condition_182 in pn_value:
            self.table_condition = "S8E2-4K139-AAVB"
        elif self.pn_condition_183 in pn_value:
            self.table_condition = "S8E2-3N128-AAVB"
        elif self.pn_condition_184 in pn_value:
            self.table_condition = "AY5A74AC101"
        elif self.pn_condition_185 in pn_value:
            self.table_condition = "P32339234"
        elif self.pn_condition_186 in pn_value:
            self.table_condition = "19103304100"
        elif self.pn_condition_187 in pn_value:
            self.table_condition = "AZ5A74AB801"
        elif self.pn_condition_188 in pn_value:
            self.table_condition = "AZ5A74AC201"
        elif self.pn_condition_189 in pn_value:
            self.table_condition = "19103303900"
        elif self.pn_condition_190 in pn_value:
            self.table_condition = "16GKU9760"
        elif self.pn_condition_191 in pn_value:
            self.table_condition = "#7T0407271B"
        elif self.pn_condition_192 in pn_value:
            self.table_condition = "14633300802"
        elif self.pn_condition_193 in pn_value:
            self.table_condition = "19103303800"
        elif self.pn_condition_194 in pn_value:
            self.table_condition = "11773505401"
        elif self.pn_condition_195 in pn_value:
            self.table_condition = "DW93-3N129-AC23"
        elif self.pn_condition_196 in pn_value:
            self.table_condition = "19103303500"
        elif self.pn_condition_197 in pn_value:
            self.table_condition = "HAT1743305400"
        elif self.pn_condition_198 in pn_value:
            self.table_condition = "16GKU9L20"
        elif self.pn_condition_199 in pn_value:
            self.table_condition = "1174330P037"
        elif self.pn_condition_200 in pn_value:
            self.table_condition = "110519223"
        elif self.pn_condition_201 in pn_value:
            self.table_condition = "11743305500"
        elif self.pn_condition_202 in pn_value:
            self.table_condition = "AZ5A74AA601"
        elif self.pn_condition_203 in pn_value:
            self.table_condition = "AY5A74AA501"
        elif self.pn_condition_204 in pn_value:
            self.table_condition = "9W83-4K139-EE"
        elif self.pn_condition_205 in pn_value:
            self.table_condition = "11773303903"
        elif self.pn_condition_206 in pn_value:
            self.table_condition = "#971501201AF"
        elif self.pn_condition_207 in pn_value:
            self.table_condition = "S8E2-3N128-AA"
        elif self.pn_condition_208 in pn_value:
            self.table_condition = "DW93-4K138-AD"
        elif self.pn_condition_209 in pn_value:
           self. table_condition = "GX53-3N129-AC23"
        elif self.pn_condition_210 in pn_value:
            self.table_condition = "GX73-4K138-DC"
        elif self.pn_condition_211 in pn_value:
            self.table_condition = "11773304903"
        elif self.pn_condition_212 in pn_value:
            self.table_condition = "11773303703"
        elif self.pn_condition_213 in pn_value:
            self.table_condition = "1174350R006"
        elif self.pn_condition_214 in pn_value:
            self.table_condition = "11783308800"
        elif self.pn_condition_215 in pn_value:
            self.table_condition = "#9Y4501201A"
        elif self.pn_condition_216 in pn_value:
            self.table_condition = "#9Y4501201B"
        elif self.pn_condition_217 in pn_value:
            self.table_condition = "11783308900"
        elif self.pn_condition_218 in pn_value:
            self.table_condition = "#4MN407271B"
        elif self.pn_condition_219 in pn_value:
            self.table_condition = "T3A3-3N129-AA"
        elif self.pn_condition_220 in pn_value:
            self.table_condition = "T3A3-3N128-AA"
        elif self.pn_condition_221 in pn_value:
            self.table_condition = "505602770028203"
        elif self.pn_condition_222 in pn_value:
            self.table_condition = "16GKU9G30"
        elif self.pn_condition_223 in pn_value:
            self.table_condition = "T3A3-3N128-ABPT"
        elif self.pn_condition_224 in pn_value:
            self.table_condition = "T3A3-3N129-ABPT"
        elif self.pn_condition_225 in pn_value:
            self.table_condition = "16GKU9790"
        elif self.pn_condition_226 in pn_value:
            self.table_condition = "10466771"
        elif self.pn_condition_227 in pn_value:
            self.table_condition = "#4M0407271S"
        elif self.pn_condition_228 in pn_value:
            self.table_condition = "[)>06P32324455T"
        elif self.pn_condition_229 in pn_value:
            self.table_condition = "1177330R000"
        elif self.pn_condition_230 in pn_value:
            self.table_condition = "670158933"
        elif self.pn_condition_231 in pn_value:
            self.table_condition = "19103304300"
        elif self.pn_condition_232 in pn_value:
            self.table_condition = "#2N0407272AA"
        elif self.pn_condition_233 in pn_value:
            self.table_condition = "10476052"
        elif self.pn_condition_234 in pn_value:
            self.table_condition = "10462062"
        elif self.pn_condition_235 in pn_value:
            self.table_condition = "[)>06P32339226T"
        elif self.pn_condition_236 in pn_value:
            self.table_condition = "46354802"
        elif self.pn_condition_237 in pn_value:
            self.table_condition = "#2N0407272AB"
        elif self.pn_condition_238 in pn_value:
            self.table_condition = "AZ5A74AB201"
        elif self.pn_condition_239 in pn_value:
            self.table_condition = "AY5A74AB101"
        elif self.pn_condition_240 in pn_value:
            self.table_condition = "#4MN407271"
        elif self.pn_condition_241 in pn_value:
            self.table_condition = "#85E407271AE"
        elif self.pn_condition_242 in pn_value:
            self.table_condition = "AZ5A74AC801"
        elif self.pn_condition_243 in pn_value:
            self.table_condition = "AY5A74AC701"
        elif self.pn_condition_244 in pn_value:
            self.table_condition = "11743305700"
        elif self.pn_condition_245 in pn_value:
            self.table_condition = "11783503600"
        elif self.pn_condition_246 in pn_value:
            self.table_condition = "#9Y4501201E"
        elif self.pn_condition_247 in pn_value:
            self.table_condition = "T3A3-4B402-ABPT"
        elif self.pn_condition_248 in pn_value:
            self.table_condition = "16GKU9690"
        elif self.pn_condition_249 in pn_value:
            self.table_condition = "11783503500"
        elif self.pn_condition_250 in pn_value:
            self.table_condition = "#9Y4501201D"
        elif self.pn_condition_251 in pn_value:
            self.table_condition = "AZ5A74AA201"
        elif self.pn_condition_252 in pn_value:
            self.table_condition = "01380047080"
        elif self.pn_condition_253 in pn_value:
            self.table_condition = "01380050080"
        elif self.pn_condition_254 in pn_value:
            self.table_condition = "19103303700"
        elif self.pn_condition_255 in pn_value:
            self.table_condition = "AZ5A74AA801"
        elif self.pn_condition_256 in pn_value:
            self.table_condition = "AY5A74AA701"
        elif self.pn_condition_257 in pn_value:
            self.table_condition = "11783500000"
        elif self.pn_condition_258 in pn_value:
            self.table_condition = "11783500100"
        elif self.pn_condition_259 in pn_value:
            self.table_condition = "#85E407271AK"
        elif self.pn_condition_260 in pn_value:
            self.table_condition = "11773302701"
        elif self.pn_condition_261 in pn_value:
            self.table_condition = "J9C3-3B436-AE"
        elif self.pn_condition_262 in pn_value:
            self.table_condition = "#4M0407271T"
        elif self.pn_condition_263 in pn_value:
            self.table_condition = "#4MN407271C"
        elif self.pn_condition_264 in pn_value:
            self.table_condition = "16GKU9F20"
        elif self.pn_condition_265 in pn_value:
            self.table_condition = "11743305400"
        elif self.pn_condition_266 in pn_value:
            self.table_condition = "10466758"
        elif self.pn_condition_267 in pn_value:
            self.table_condition = "463568890"
        elif self.pn_condition_268 in pn_value:
            self.table_condition = "463568900"
        elif self.pn_condition_269 in pn_value:
            self.table_condition = "463548020"
        elif self.pn_condition_270 in pn_value:
            self.table_condition = "#983501202D"
        elif self.pn_condition_271 in pn_value:
            self.table_condition = "105150130"
        elif self.pn_condition_272 in pn_value:
            self.table_condition = "#95C501201S"
        elif self.pn_condition_273 in pn_value:
            self.table_condition = "1380048080"
        elif self.pn_condition_274 in pn_value:
            self.table_condition = "01380045080"
        elif self.pn_condition_275 in pn_value:
            self.table_condition = "#4M0501201G"
        elif self.pn_condition_276 in pn_value:
            self.table_condition = "CPLA-4K139-EA"
        elif self.pn_condition_277 in pn_value:
            self.table_condition = "CPLA-4K138-CA"
        elif self.pn_condition_278 in pn_value:
            self.table_condition = "MY42-3D402-AA"
        elif self.pn_condition_279 in pn_value:
            self.table_condition = "M8B2-3D402-AA"
        elif self.pn_condition_280 in pn_value:
            self.table_condition = "GX53-3N128-AC"
        elif self.pn_condition_281 in pn_value:
            self.table_condition = "M8E2-3D402-BB"
        elif self.pn_condition_282 in pn_value:
            self.table_condition = "1236350R001"
        else:
            self.table_condition = None

        if self.table_condition is not None:
            all_entries_valid = self.check_table_entries()
            numer_pojemnika_valid = self.validate_numer_pojemnika()

            if all_entries_valid and numer_pojemnika_valid:
                all_windows_green = all(entry.cget("background") == 'green' for entry in self.table_entries)
                any_windows_orange = any(entry.cget("background") == 'orange' for entry in self.table_entries)

                if all_windows_green and not any_windows_orange:
                    self.button_verify["state"] = "normal"
                    self.button_add["state"] = "normal"
                else:
                    self.button_verify["state"] = "normal"
                    self.button_add["state"] = "disabled"
                self.master.focus_set()
                self.master.focus_force()

            else:
                self.button_verify["state"] = "normal"
                self.button_add["state"] = "disabled"
        else:
            messagebox.showwarning("Ostrzeżenie", "Niepoprawna wartość w rubryce PN!")

    def check_table_entries(self):
        all_entries_valid = True
        seen_values = set()

        for i, entry in enumerate(self.table_entries):
            cell_value = entry.get()
            cell_color = entry.cget("background")

            if not cell_value:
                entry.configure(bg='red')
                all_entries_valid = False
            elif cell_value in seen_values:
                entry.configure(bg='orange')
                all_entries_valid = False
            elif self.table_condition is not None and not cell_value.startswith(self.table_condition):
                entry.configure(bg='red')
            elif self.table_condition in cell_value:
                entry.configure(bg='green')
            else:
                entry.configure(bg='red')

            seen_values.add(cell_value)
            self.original_table_values[i] = cell_value

        return all_entries_valid

    def validate_numer_pojemnika(self, event=None):
        numer_pojemnika_value = self.entry_container.get()
        if not numer_pojemnika_value or len(numer_pojemnika_value) != 9:
            messagebox.showwarning("Ostrzeżenie", "Numer pojemnika powinien zawierać dokładnie 9 znaków!")
            self.entry_container.delete(0, tk.END)
            return False
        else:
            return True

    def weryfikuj_dane(self, prefix, entry, rubryka):
        if not entry.get().startswith(prefix):
            messagebox.showwarning("Ostrzeżenie", f"Niepoprawne dane w rubryce {rubryka}!")
            entry.delete(0, tk.END)

    def dodaj_dane(self):
        for entry in self.table_entries:
            if entry.cget("background") == 'yellow':
                entry.configure(bg='yellow')  # Wartość powtórzona
                self.button_add["state"] = "disabled"
                return

        self.weryfikuj_dane("R", self.entry_container, "Numer pojemnika")
        self.weryfikuj_dane("30S", self.entry_pn, "PN")

        pn = self.entry_pn.get()
        numer_pojemnika = self.entry_container.get()
        ilosc_sztuk = self.entry_quantity.get()
        powod_nok = self.entry_powod_nok.get()
        dzialania = self.slider_dzialania.get()

        data = time.strftime('%Y-%m-%d')
        godzina = time.strftime('%H:%M:%S')

        if not pn or not numer_pojemnika or not ilosc_sztuk or not powod_nok or not dzialania:
            messagebox.showwarning("Ostrzeżenie", "Wypełnij wszystkie pola!")
            return

        self.dane = {"PN": [], "Numer pojemnika": [], "Ilość sztuk": [], "Powód sztuki NOK": [],
                     "Data": [], "Godzina": [], "Działania": [], "Tabela": []}
        self.dane["PN"].append(pn)
        self.dane["Numer pojemnika"].append(numer_pojemnika)
        self.dane["Ilość sztuk"].append(ilosc_sztuk)
        self.dane["Powód sztuki NOK"].append(powod_nok)
        self.dane["Działania"].append(dzialania)
        self.dane["Data"].append(data)
        self.dane["Godzina"].append(godzina)
        tabela = [entry.get() for entry in self.table_entries]
        self.dane["Tabela"].append(tabela)

        self.generuj_raport()

        self.button_add["state"] = "disabled"
        self.button_verify["state"] = "normal"

        self.entry_pn.delete(0, tk.END)
        self.entry_container.delete(0, tk.END)
        self.entry_quantity.delete(0, tk.END)
        self.entry_powod_nok.delete(0, tk.END)

        for entry in self.table_entries:
            entry.delete(0, tk.END)

        self.label_quantity.config(text="Ilość sztuk NOK:")

        if not self.dane["PN"]:
            messagebox.showwarning("Ostrzeżenie", "Brak danych do wygenerowania raportu!")
            return

        file_path = os.path.join(os.path.dirname(__file__), "spis.xlsx")
        if os.path.exists(file_path):
            existing_df = pd.read_excel(file_path)
            self.dane_df = pd.DataFrame(self.dane)
            combined_df = pd.concat([existing_df, self.dane_df], ignore_index=True)
        else:
            combined_df = pd.DataFrame(self.dane)

        combined_df.to_excel(file_path, index=False)
        messagebox.showinfo("Sukces", "Raport zaktualizowany i zapisany jako spis.xlsx")

        self.slider_dzialania.set("")

        self.generuj_raport()

        self.button_add["state"] = "disabled"
        self.button_verify["state"] = "normal"

        self.entry_pn.delete(0, tk.END)
        self.entry_container.delete(0, tk.END)
        self.entry_quantity.delete(0, tk.END)
        self.entry_powod_nok.delete(0, tk.END)

        for entry in self.table_entries:
            entry.delete(0, tk.END)

        self.label_quantity.config(text="Ilość sztuk NOK:")

        if not self.dane["PN"]:
            messagebox.showwarning("Ostrzeżenie", "Brak danych do wygenerowania raportu!")
            return

        file_path = os.path.join(os.path.dirname(__file__), "spis.xlsx")
        if os.path.exists(file_path):
            existing_df = pd.read_excel(file_path)
            self.dane_df = pd.DataFrame(self.dane)
            combined_df = pd.concat([existing_df, self.dane_df], ignore_index=True)
        else:
            combined_df = pd.DataFrame(self.dane)

        combined_df.to_excel(file_path, index=False)
        messagebox.showinfo("Sukces", "Raport zaktualizowany i zapisany jako spis.xlsx")

        self.slider_dzialania.set("")

    def generuj_raport(self):
        pass

    def otworz_link(self):
        import webbrowser
        webbrowser.open("https://www.youtube.com/watch?v=3dHpEfmegOA")

    def update_datetime(self):
        current_date = time.strftime('%Y-%m-%d')
        current_time = time.strftime('%H:%M:%S')
        self.label_current_date.config(text=current_date)
        self.label_current_time.config(text=current_time)
        self.master.after(1000, self.update_datetime)

    def on_closing(self):
        if messagebox.askokcancel("Zamknij program", "Czy na pewno chcesz zamknąć program?"):
            self.master.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    program = ProgramGUI(root)
    root.mainloop()