import tkinter as tk
from tkinter import messagebox
import requests
from datetime import datetime
import openpyxl
import os

API_KEY = "61c18defa01913e8ebb7d89753e70a01"
CIDADE = "Sao Paulo,BR"
NOME_ARQUIVO_EXCEL = "previsao_tempo_sp.xlsx"
COLUNAS = ["Data / Hora", "Temperatura", "Umidade do Ar"]


def fetch_weather_data_api():

    url_api = f"https://api.openweathermap.org/data/2.5/weather?q={CIDADE}&appid={API_KEY}&units=metric&lang=pt_br"
    
    try:
        response = requests.get(url_api)
        response.raise_for_status()
        data = response.json()
        
        temperatura = data['main']['temp']
        umidade = data['main']['humidity']
        
        temperatura_formatada = f"{temperatura}°C"
        umidade_formatada = f"{umidade}%"
        
        return temperatura_formatada, umidade_formatada

    except requests.exceptions.HTTPError as e:
        if e.response.status_code == 401:
             messagebox.showerror("Erro de API", "Chave de API inválida ou não ativada. Verifique a chave ou aguarde alguns minutos para ela ser ativada.")
        else:
            messagebox.showerror("Erro de API", f"Não foi possível buscar os dados.\nErro: {e}")
        return None, None
    except requests.exceptions.RequestException as e:
        messagebox.showerror("Erro de Conexão", f"Não foi possível conectar à API.\nVerifique sua conexão.\n\nDetalhes: {e}")
        return None, None
    except KeyError:
        messagebox.showerror("Erro de Dados", "A resposta da API não continha os dados esperados.")
        return None, None


def save_to_spreadsheet(data):
    try:
        if not os.path.exists(NOME_ARQUIVO_EXCEL):
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Previsões"
            sheet.append(COLUNAS)
        else:
            workbook = openpyxl.load_workbook(NOME_ARQUIVO_EXCEL)
            sheet = workbook.active
        
        sheet.append(data)
        workbook.save(NOME_ARQUIVO_EXCEL)
        return True
    except Exception as e:
        messagebox.showerror("Erro de Gravação", f"Não foi possível salvar os dados na planilha.\n\nDetalhes: {e}")
        return False

def main_action():
    print("Buscando previsão pela API...")
    temperatura, umidade = fetch_weather_data_api()
    
    if temperatura is None or umidade is None:
        return

    data_hora_atual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    dados_para_salvar = [data_hora_atual, temperatura, umidade]
    
    if save_to_spreadsheet(dados_para_salvar):
        print(f"Dados salvos com sucesso em '{NOME_ARQUIVO_EXCEL}'")
        messagebox.showinfo("Sucesso", f"Previsão atualizada com sucesso!\n\n- Data/Hora: {data_hora_atual}\n- Temperatura: {temperatura}\n- Umidade: {umidade}")

def create_gui():
    window = tk.Tk()
    window.title("Previsão do tempo de São Paulo")
    window.geometry("400x150")
    window.resizable(False, False)

    frame = tk.Frame(window)
    frame.pack(expand=True)

    label = tk.Label(frame, text="Atualizar previsão na planilha:", font=("Segoe UI", 12))
    label.pack(pady=(0, 10))

    button = tk.Button(frame, text="Buscar previsão", font=("Segoe UI", 11, "bold"), command=main_action, width=20, height=2)
    button.pack(pady=(0, 10))
    
    window.mainloop()

if __name__ == "__main__":
    create_gui()