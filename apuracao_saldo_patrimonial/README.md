# Apuração de Saldo Patrimonial

Sistema que processa dois arquivos Excel (998 e 990), filtra contas por tipo e atributo, e preenche a planilha **Apuração de saldo patrimonial** com os valores apurados.

## Arquivos

- **998**: coluna A = conta contábil, G = débito atual, H = crédito atual.
- **990**: coluna A = conta contábil, D = tipo de conta, J = atributo da conta (F ou P).
- **Apuração de saldo patrimonial**: planilha modelo onde os valores são gravados nas células definidas pelas regras.

## Regras

1. No **990** são consideradas apenas contas com **tipo = 5** e **atributo (coluna J) = F ou P**.
2. As contas do **998** são cruzadas com as do **990** (coluna A). Para cada conta que passar no filtro do 990, o valor é obtido no 998 (coluna G ou H; usa-se o saldo débito − crédito).
3. O resultado é somado por grupo e escrito na planilha de apuração conforme o mapeamento abaixo.

### Mapeamento (coluna:linha → critério)


| Célula | Contas (prefixo ou exata)   | Atributo 990 |
| ------ | --------------------------- | ------------ |
| B5     | começam com 111             | F            |
| B6     | começam com 113             | F            |
| B7     | começam com 114             | F            |
| B8     | começam com 12              | F            |
| B11    | começam com 11              | P            |
| B12    | começam com 12              | P            |
| B15    | começam com 21              | F            |
| B16    | começam com 22              | F            |
| B17    | conta exata 631100000000000 | (tipo 5)     |
| B18    | conta exata 631710000000000 | (tipo 5)     |
| B21    | começam com 21              | P            |
| B22    | começam com 22              | P            |


## Uso

Todo o código está em **`apuracao_saldo_patrimonial.py`** (lógica + interface).

### Interface Streamlit

```bash
pip install -r requirements.txt
streamlit run apuracao_saldo_patrimonial.py
```

Envie os três arquivos .xlsx e clique em **Processar e gerar apuração**. Em seguida, baixe a planilha preenchida.

### Como módulo

```python
from apuracao_saldo_patrimonial import main

main(
    arquivo_998="caminho/998.xlsx",
    arquivo_990="caminho/990.xlsx",
    arquivo_apuracao_template="caminho/apuracao_modelo.xlsx",
    arquivo_apuracao_saida="caminho/apuracao_preenchida.xlsx",
)
```

Também é possível passar objetos file-like (por exemplo, `BytesIO` ou arquivos enviados por upload) em vez de caminhos.