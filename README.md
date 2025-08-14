# Consulta de Faturamento Anual de Empresas

Este projeto é uma aplicação Python com interface gráfica para consulta rápida de faturamento anual de empresas a partir de arquivos Excel.  
Foi desenvolvido com **Tkinter** e **ttkbootstrap** para oferecer uma interface amigável e moderna.

---

## ✨ Funcionalidades
- Carrega um arquivo Excel contendo múltiplas abas de faturamento.
- Concatena todos os dados em uma única visualização.
- Busca empresas pelo nome (insensível a maiúsculas/minúsculas).
- Formata valores numéricos como moeda (R$).
- Oculta automaticamente as duas últimas colunas da planilha na exibição (ex.: `MÉTRICA` e `VALOR`).
- Interface com tema visual agradável (`minty`).

---

## 📷 Exemplo de uso
1. Abra o programa.
2. Clique em **"Escolher arquivo Excel"** e selecione o arquivo.
3. Digite o nome (ou parte do nome) da empresa no campo de busca.
4. Clique em **"Buscar"** e veja o resultado formatado na tabela.

---

## 🛠 Tecnologias utilizadas
- **Python 3.12+**
- [pandas](https://pandas.pydata.org/)
- [tkinter](https://docs.python.org/3/library/tkinter.html)
- [ttkbootstrap](https://ttkbootstrap.readthedocs.io/en/latest/)

---

## 📦 Instalação
Clone o repositório e instale as dependências:

```bash
git clone https://github.com/seu-usuario/consulta-faturamento.git
cd consulta-faturamento
pip install pandas ttkbootstrap openpyxl

---
````
## ▶️ Como executar
python app.py


Substitua app.py pelo nome do arquivo onde o código está salvo.

---

## 📂 Estrutura do projeto
consulta-faturamento/
│
├── app.py          # Código principal da aplicação
├── README.md       # Este arquivo
└── requirements.txt (opcional)

---

## ⚠️ Observações

A planilha deve conter uma coluna chamada CLIENTE.

As duas últimas colunas do Excel serão ocultadas automaticamente na exibição da tabela.

Para ler arquivos Excel, é necessário ter o pacote openpyxl instalado.
