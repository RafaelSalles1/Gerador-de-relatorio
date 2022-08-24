from tkinter import *
from tkinter import PhotoImage
import tkinter as tk
import PIL
from tkinter import filedialog
from PIL import Image
from PIL import ImageTk
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.section import *
from docx.enum.text import *
from docx.enum.style import *
from docx.enum.dml import *


class MainApp():
	def __init__(self, root):

		self.root = root

		self.imagem_disparco = Image.open("X:\LOGO DISPARCO\PNG-transparente.png")
		self.imagem_disparco = self.imagem_disparco.resize((280, 130))
		self.logo_disparco = ImageTk.PhotoImage(self.imagem_disparco)

		self.canvas2 = tk.Canvas(root, width = 900, height = 501, bg = "#D8D0C8", confine = TRUE)
		self.canvas2.pack()
		self.canvas2.create_image(150, 100, image = self.logo_disparco)

		self.num_proc = tk.Entry(self.canvas2)
		self.num_proc.place(width = 100, height = 30, x = 700, y = 100)

		self.nome_cliente = tk.Entry(self.canvas2)
		self.nome_cliente.place(width = 300, height = 30, x = 500, y = 50)

		self.frame1 = tk.Frame(self.canvas2, bg = "#D8D0C8")
		self.frame1.place(relwidth = 0.33, relheight = 0.66, x = 0, y = 167)

		self.frame2 = tk.Frame(self.canvas2, bg = "#D8D0C8")
		self.frame2.place(relwidth = 0.33, relheight = 0.66, x = 301, y = 167)

		self.frame3 = tk.Frame(self.canvas2, bg = "#D8D0C8")
		self.frame3.place(relwidth = 0.33, relheight = 0.66, x = 601, y = 167)

		self.label_nome_cliente = tk.Label(self.canvas2, text = "Nome do cliente:", bg = "#D8D0C8", fg = "black", font = ("helvetica", 9, "bold"))
		self.label_nome_cliente.place(width = 100, height = 30, x = 400, y = 50)

		self.label_num_proc = tk.Label(self.canvas2, text = "Numero do processo:", bg = "#D8D0C8", fg = "black", font = ("helvetica", 9, 'bold'))
		self.label_num_proc.place(width = 128, height = 30, x = 400, y = 100)

		self.generate_button = tk.Button(self.frame3, text = "Gerar Relatório", bg = "#D8D0C8", fg = "black", command = self.click)
		self.generate_button.place(relwidth = 0.5, relheight = 0.1, x = 120,  y = 270)

		self.var_introducao = tk.BooleanVar()
		self.var_melhorias = tk.IntVar()
		self.var_geracao_de_vapor = tk.IntVar()
		self.var_distribuicao_de_vapor = tk.IntVar()
		self.var_boas_praticas = tk.IntVar()
		self.var_adeq_in_out = tk.IntVar()
		self.var_adeq_linhas_ext = tk.IntVar()
		self.var_desaer_rede = tk.IntVar()
		self.var_dren_coletiva = tk.IntVar()
		self.var_vapor_preso_teoria = tk.IntVar()
		self.var_vapor_preso_cilindros = tk.IntVar()
		self.var_tubos_isolamento = tk.IntVar()
		self.var_purgadores_termo_x_boia = tk.IntVar()
		self.var_locais_vazamentos_ext = tk.IntVar()
		self.var_valvulas_bypass = tk.IntVar()

		self.cb_introducao = tk.Checkbutton(self.frame1, text="Introdução", bg = "#D8D0C8", variable=self.var_introducao, pady = 5, padx = 0)
		self.cb_introducao.place(relx = 0, rely = 0)

		self.cb_melhorias = tk.Checkbutton(self.frame1, text="Melhorias", bg = "#D8D0C8", variable=self.var_melhorias, pady = 5)
		self.cb_melhorias.place(relx = 0, rely = 0.15)

		self.cb_geracao_de_vapor = tk.Checkbutton(self.frame1, text="Geração de Vapor (teoria)", bg = "#D8D0C8", variable=self.var_geracao_de_vapor, pady = 5)
		self.cb_geracao_de_vapor.place(relx = 0, rely = 0.30)

		self.cb_distribuicao_de_vapor = tk.Checkbutton(self.frame1, text="Distribuição de Vapor", bg = "#D8D0C8",variable=self.var_distribuicao_de_vapor, pady = 5, command = self.check_status_distribuicao)
		self.cb_distribuicao_de_vapor.place(relx = 0, rely = 0.45)

		self.cb_boas_praticas = tk.Checkbutton(self.frame1, text="Boas praticas para remoção de condensado", bg = "#D8D0C8", variable=self.var_boas_praticas, pady = 5)
		self.cb_boas_praticas.place(relx = 0, rely = 0.60)

		self.cb_adeq_in_out = tk.Checkbutton(self.frame2, text="Adequação necessidade dren. Inlet/Outlet", bg = "#D8D0C8",variable=self.var_adeq_in_out, pady = 5)
		self.cb_adeq_in_out.place(relx = 0, rely = 0)

		self.cb_adeq_linhas_ext = tk.Checkbutton(self.frame2, text="Adequação drenagens de linha existentes", bg = "#D8D0C8",variable=self.var_adeq_linhas_ext, pady = 5)
		self.cb_adeq_linhas_ext.place(relx = 0, rely = 0.15)

		self.cb_desaer_rede = tk.Checkbutton(self.frame2, text="Desaeração da rede de vapor", bg ="#D8D0C8", variable=self.var_desaer_rede, pady = 5)
		self.cb_desaer_rede.place(relx = 0, rely = 0.3)

		self.cb_dren_coletiva = tk.Checkbutton(self.frame2, text="Drenagem coletiva", bg = "#D8D0C8",variable=self.var_dren_coletiva, pady = 5)
		self.cb_dren_coletiva.place(relx = 0, rely = 0.45)

		self.cb_vapor_preso_teoria = tk.Checkbutton(self.frame2, text="Vapor Preso (Teoria)", bg = "#D8D0C8", variable=self.var_vapor_preso_teoria, pady = 5)
		self.cb_vapor_preso_teoria.place(relx = 0, rely = 0.6)

		self.cb_vapor_preso_cilindros = tk.Checkbutton(self.frame3, text="Vapor Preso (Cilindros)", bg = "#D8D0C8", variable=self.var_vapor_preso_cilindros, pady = 5)
		self.cb_vapor_preso_cilindros.place(relx = 0, rely = 0)

		self.cb_tubos_isolamento = tk.Checkbutton(self.frame3, text="Tubulações sem isolamento", bg = "#D8D0C8", variable=self.var_tubos_isolamento, pady = 5)
		self.cb_tubos_isolamento.place(relx = 0, rely = 0.15)

		self.cb_purgadores_termo_x_boia = tk.Checkbutton(self.frame3, text="Purgadores Termo x Boias", bg = "#D8D0C8", variable=self.var_purgadores_termo_x_boia, pady = 5)
		self.cb_purgadores_termo_x_boia.place(relx = 0, rely = 0.3)

		self.cb_locais_vazamentos_ext = tk.Checkbutton(self.frame3, text="Locais com vaz. externos", bg = "#D8D0C8",variable=self.var_locais_vazamentos_ext, pady = 5)
		self.cb_locais_vazamentos_ext.place(relx = 0, rely = 0.45)

		self.cb_valvulas_bypass = tk.Checkbutton(self.frame3, text="Valvulas By-Pass", bg = "#D8D0C8", variable=self.var_valvulas_bypass, pady = 5)
		self.cb_valvulas_bypass.place(relx = 0, rely = 0.6)


	def click(self):
		self.nome_cliente_str = self.nome_cliente.get()
		self.num_proc_str = self.num_proc.get()
		self.save_name = "RG-" + self.num_proc_str + "-001-20-00-Relatório de Gerenciamento-" + self.nome_cliente_str
	

	def check_status_distribuicao(self):
		if self.var_distribuicao_de_vapor.get():
			self.caminho_distribuicao = filedialog.askopenfilenames()
			self.fotos_distribuicao_de_vapor = list(self.caminho_distribuicao)
		return self.fotos_distribuicao_de_vapor



if __name__ == "__main__":
	app = tk.Tk()
	MainApp(app)	
	app.mainloop()

















