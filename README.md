# 📊 Comparador de Planilhas

Este projeto é uma aplicação **web interativa em Streamlit** para comparar uma ou várias planilhas do Excel (`.xlsx`) e identificar similaridades e diferenças entre colunas selecionadas.

A aplicação permite:

* Upload de uma ou mais planilhas em formato `.xlsx`.
* Seleção da aba (sheet) de cada planilha para comparação.
* Visualização das planilhas carregadas diretamente na interface.
* Escolha de colunas específicas para comparar em cada planilha.
* Diferentes modos de comparação:

  * Comparação dentro da própria planilha (caso apenas uma seja carregada).
  * Primeira planilha contra as outras.
  * Comparação cruzada entre todas as planilhas.
* Comparação baseada em similaridade de strings (`difflib`).
* Barra de progresso mostrando quantos itens já foram processados, total e estimativa de tempo restante.
* Geração de um relatório final em Excel formatado com:

  * Coluna original.
  * Item mais similar encontrado nas outras planilhas.
  * Diferenças detectadas.
* Download da planilha de resultados com formatação visual aplicada.

---

## 🚀 Tecnologias Utilizadas

* [Python 3.9+](https://www.python.org/)
* [Streamlit](https://streamlit.io/) (interface web)
* [Pandas](https://pandas.pydata.org/) (manipulação de planilhas)
* [OpenPyXL](https://openpyxl.readthedocs.io/) (formatação e exportação para Excel)
* [Difflib](https://docs.python.org/3/library/difflib.html) (comparação de similaridade de strings)

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

1. Preencha o formulário inicial informando:

   * Quantidade de planilhas.
   * Quantidade de colunas a serem comparadas.
   * Modo de comparação (se houver mais de uma planilha).
2. Faça upload das planilhas.
3. Escolha a aba de cada planilha.
4. Visualize os dados antes de iniciar a comparação.
5. Escolha as colunas que deseja comparar em cada planilha.
6. Clique em **"Iniciar Comparação"**.
7. Acompanhe a barra de progresso com itens processados e estimativa de tempo.
8. Baixe a planilha de resultados formatada.

---

## 📥 Saída

A planilha final gerada contém:

* **Coluna original** de cada planilha.
* **Item mais similar encontrado** nas outras planilhas.
* **Diferenças** (caso existam).

O arquivo de saída é salvo como:

```
Diferenças.xlsx
```

Com formatação aplicada:

* Cabeçalhos em negrito e fundo cinza.
* Largura de colunas ajustada.
* Texto quebrado para células longas.

---

## 📌 Melhorias Futuras

* Comparação de múltiplas colunas por aba simultaneamente.
* Ajustar sensibilidade da comparação de similaridade.
* Destacar visualmente células com diferenças detectadas.
* Interface com opções avançadas de filtros.
* Exportação em outros formatos (CSV, JSON).

---

## 📝 Licença

Este projeto é distribuído sob a licença MIT.
Sinta-se à vontade para usar, modificar e compartilhar.

---

## 👨‍💻 Autor

Desenvolvido por **Nathanael Secundo Cardoso** 🎯
Estudante de Ciência da Computação | Apaixonado por automação e análise de dados.