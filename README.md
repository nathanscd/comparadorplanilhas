````
# 📊 Comparador de Planilhas

Este projeto é uma aplicação **web interativa em Streamlit** para comparar duas planilhas do Excel (`.xlsx`) e identificar similaridades e diferenças entre colunas selecionadas.  

A aplicação permite:

- Upload de duas planilhas em formato `.xlsx`.
- Seleção da aba (sheet) de cada planilha para comparação.
- Visualização das planilhas carregadas diretamente na interface.
- Escolha de colunas específicas para comparar.
- Comparação baseada em similaridade de strings (`difflib`).
- Geração de um relatório final em Excel formatado com:
  - Coluna original.
  - Item mais similar encontrado na outra planilha.
  - Diferenças detectadas.
- Download da planilha de resultados.
- Barra de progresso mostrando o andamento da comparação.

---

## 🚀 Tecnologias Utilizadas

- [Python 3.9+](https://www.python.org/)
- [Streamlit](https://streamlit.io/) (interface web)
- [Pandas](https://pandas.pydata.org/) (manipulação de planilhas)
- [OpenPyXL](https://openpyxl.readthedocs.io/) (formatação e exportação para Excel)
- [Difflib](https://docs.python.org/3/library/difflib.html) (comparação de similaridade de strings)

---

## 📦 Instalação e Uso

1. Clone este repositório:
  ```
   git clone https://github.com/seu-usuario/comparador-planilhas.git
   cd comparador-planilhas
  ```

2. Crie e ative um ambiente virtual (opcional, mas recomendado):

   ```
   python -m venv venv
   source venv/bin/activate   # Linux/Mac
   venv\Scripts\activate      # Windows
   ```

3. Instale as dependências:

   ```
   pip install -r requirements.txt
   ```

4. Execute o aplicativo:

   ```
   streamlit run app.py
   ```

5. Abra no navegador:

   ```
   http://localhost:8501
   ```

---

## 📁 Estrutura do Projeto

```
📂 comparador-planilhas
 ├── app.py             # Código principal do Streamlit
 ├── requirements.txt   # Dependências do projeto
 ├── README.md          # Documentação
```

---

## 🖼️ Exemplo de Uso

1. Carregue duas planilhas `.xlsx`.
2. Escolha a aba de cada planilha.
3. Visualize os dados antes de comparar.
4. Escolha as colunas a serem comparadas.
5. Clique em **"Iniciar Conversão"**.
6. Acompanhe a barra de progresso.
7. Baixe a planilha de resultados.

---

## 📥 Saída

A planilha final gerada contém:

* **Coluna original (planilha 1)**.
* **Item mais similar encontrado (planilha 2)**.
* **Diferenças** (caso existam).

O arquivo de saída é salvo como:

```
Diferenças.xlsx
```

---

## 📌 Melhorias Futuras

* Suporte a comparação de múltiplas colunas.
* Ajustar sensibilidade da comparação de similaridade.
* Interface com opções avançadas de filtros.
* Exportação em outros formatos (CSV, JSON).

---

## 📝 Licença

Este projeto é distribuído sob a licença MIT.
Sinta-se à vontade para usar, modificar e compartilhar.

---

## 👨‍💻 Autor

Desenvolvido por **\Nathanael Secundo Cardoso** 🎯
Estudante de Ciência da Computação | Apaixonado por automação e análise de dados.