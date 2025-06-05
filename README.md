# Controle de Estoque

**Aplicação Desktop em Python para gerenciar estoques via arquivos Excel.**

![Python](https://img.shields.io/badge/python-3.8%2B-blue)
![PySimpleGUI](https://img.shields.io/badge/PySimpleGUI-v4-brightgreen)
![OpenPyXL](https://img.shields.io/badge/OpenPyXL-v3-orange)

---

## 📖 Sobre

Este projeto é um sistema simples de controle de estoque, desenvolvido em Python com PySimpleGUI e OpenPyXL.  
Ele permite:
- Autenticação de usuários pré-configurados.
- Visualização em tempo real do estoque via tabela.
- Entrada/saída de itens com registro automático em planilhas Excel.
- Cadastro e exclusão de equipamentos.
- Consulta histórica de entradas e retiradas.
- Logging de todas as ações (para auditoria).

---

## 🚀 Tecnologias Utilizadas

- **Python 3.8+**
- [PySimpleGUI](https://pypi.org/project/PySimpleGUI/) (GUI simplificada)
- [OpenPyXL](https://pypi.org/project/openpyxl/) (leitura/gravação de arquivos `.xlsx`)
- **Logging** para manter histórico de ações em `.txt`
- Planilhas Excel como “banco de dados”:
  - `estoque.xlsx` (base principal de quantidades)
  - `registro_entradas.xlsx` (historico de entradas)
  - `registro_retiradas.xlsx` (historico de saídas)


