# ⚡ Automação de Coleta de Dados – ISSO Digital

Este projeto realiza a automação da coleta de dados elétricos (tensão fase-neutro, fase-fase e correntes) a partir da plataforma [ISSO Digital](https://dmi.isso.digital). Os dados são exportados em CSV, tratados e inseridos automaticamente em uma planilha de controle.

---

## 📋 Funcionalidades

- Login automatizado na plataforma ISSO Digital
- Navegação até a aba de gráficos
- Fechamento de popups e sincronização de filtros (dados do dia anterior separados de forma horária)
- Exportação de três gráficos distintos:
  - Tensão fase-neutro
  - Tensão fase-fase
  - Correntes
- Tratamento dos dados CSV (conversão, filtragem, limpeza)
- Inserção dos dados em uma planilha Excel existente (definir caminho na variável `planilha_controle`)

---

## 🧰 Requisitos

- Python 3.11 ou superior
- Microsoft Edge instalado
- Driver do Microsoft Edge (`msedgedriver.exe`) compatível com sua versão do navegador

---

## 📦 Instalação

1. **Crie ou navegue até a pasta onde deseja clonar o projeto**:
   ```bash
   mkdir exemplo
   cd exemplo

2. **Clone o repositório**:
    git clone https://github.com/xSecDet/coletaDados_ISSO.git
    cd coletaDados_ISSO

3. **Instale as dependências**:
   pip install -r requirements.txt 

4. **Crie o arquivo .env com as suas credenciais**:
    Baseie-se no .env.example incluído no projeto

    IMPORTANTE MANTER A MESMA ESTRUTURA!! (SEM ESPAÇOS)
    Exemplo:
        EMAIL=seu_email@exemplo.com
        SENHA=sua_senha_segura

5. **Baixar o driver do MS Edge**:
    Coloque o msedgedriver.exe na pasta "driver", na raiz do projeto.
    Verifique se o caminho no código está correto: service = Service('./driver/msedgedriver.exe')

6. **Execute o script principal**:
    python script.py

📌 Observações
- Os dados são filtrados para o dia anterior à execução
- A planilha de destino deve estar fechada durante o processo
- A pasta "downloads" é limpa automaticamente a cada execução

---