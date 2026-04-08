# ⚡ Portal Financeiro - Scryta Contabilidade Digital

Um sistema web robusto e seguro desenvolvido para a coleta, validação e gestão de lançamentos de distribuição de lucros e dividendos. O foco principal da aplicação é garantir a conformidade fiscal (compliance) com obrigações acessórias federais (como a EFD-Reinf), automatizando a comunicação entre o escritório contábil e seus clientes.

## 🎯 O Problema que Resolvemos
Com as novas exigências da Receita Federal para a declaração mensal de distribuição de lucros, escritórios contábeis enfrentam o desafio de coletar dados precisos de seus clientes em tempo hábil. Este portal elimina o uso de planilhas e e-mails soltos, centralizando as informações em um ambiente seguro, rastreável e juridicamente válido.

## ✨ Principais Funcionalidades

### 💼 Painel do Cliente (Self-Service)
* **Gestão Multi-Empresa e Sócios:** Clientes gerenciam múltiplos CNPJs e cadastram os sócios recebedores de forma intuitiva.
* **Compartilhamento de Acesso:** Sistema de convites inteligentes que permite delegar a gestão de empresas específicas para outros e-mails corporativos.
* **Cálculo de IRRF Dinâmico (Gross-up):** O sistema alerta e calcula automaticamente a retenção na fonte caso os dividendos do sócio ultrapassem o teto de isenção (Ex: R$ 50.000,00) dentro do mês.

### 🔒 Segurança e Compliance (Não-Repúdio)
* **Assinatura Eletrônica via Token:** Nenhum lançamento entra no banco oficial sem que o usuário valide a operação através de um token OTP de 6 dígitos enviado ao seu e-mail autenticado.
* **Termo de Veracidade:** Aceite obrigatório de responsabilidade tributária e criminal antes do envio de lotes, com captura de IP, Data e Hora (Lastro Jurídico).

### 🛡️ Arquitetura Zero Trust & God Mode (Operações Internas)
* **Painel Operador (God Mode):** Administradores do escritório podem acessar todas as empresas e resolver pendências pelos clientes.
* **Trava de Domínio:** Operações internas exigem validação de identidade e só enviam tokens para e-mails restritos (`@scryta.com.br`).
* **Auditoria Transparente:** Qualquer ação feita pelo escritório ganha uma tag visual inalterável (`Op: Nome do Colaborador`), visível tanto para a gestão quanto para o cliente.

### 📊 Painel Administrativo (Gestão Contábil)
* **Visão Consolidada:** Dashboard gerencial com totalizadores de envios por Mês/Referência e Empresa.
* **Sistema de Alertas:** Identificação automática de envios retroativos (fora do prazo) e solicitações de cancelamento.
* **Exportação Dinâmica:** Geração de relatórios em Excel (`.xlsx`) formatados e prontos para integração com sistemas contábeis (Domínio/Thomson Reuters, etc.), respeitando filtros de busca, status e mês.

---

## 🛠️ Tecnologias Utilizadas

* **Backend:** Python 3, Flask
* **Banco de Dados:** MongoDB (via PyMongo)
* **Frontend:** HTML5, Tailwind CSS (Utility-first), JavaScript (Vanilla)
* **Autenticação & Segurança:** Flask-Login, Werkzeug (Hashing), ItsDangerous (Geração de Tokens)
* **E-mail (SMTP):** Flask-Mail
* **Manipulação de Dados:** OpenPyXL (Excel), Regex (Validação estrita de CPF/CNPJ)

---

## 🚀 Como executar o projeto localmente

### 1. Clonar o repositório
```bash
git clone [https://github.com/SEU-USUARIO/SEU-REPOSITORIO.git](https://github.com/SEU-USUARIO/SEU-REPOSITORIO.git)
cd SEU-REPOSITORIO
```

### 2. Criar e ativar o ambiente virtual (Recomendado)
```bash
python -m venv venv

# No Windows:
venv\Scripts\activate

# No Linux/Mac:
source venv/bin/activate
```

### 3. Instalar as dependências
```bash
pip install -r requirements.txt
```

### 4. Configurar Variáveis de Ambiente
Crie um arquivo `.env` na raiz do projeto (este arquivo é ignorado pelo Git por segurança) e adicione as seguintes chaves:

```env
SECRET_KEY=sua_chave_secreta_super_segura
MONGO_URI=mongodb+srv://usuario:senha@cluster.mongodb.net/nome_do_banco
MAIL_SERVER=smtp.seudominio.com
MAIL_PORT=465
MAIL_USE_SSL=True
MAIL_USERNAME=seu_email@dominio.com
MAIL_PASSWORD=sua_senha_de_app
MAIL_DEFAULT_SENDER=seu_email@dominio.com
```

### 5. Executar a Aplicação
```bash
python app.py
```
O sistema estará rodando em `http://localhost:5894` (ou na porta configurada no seu ambiente).

---
*Projeto desenvolvido para otimização de rotinas contábeis e garantia de conformidade fiscal.*

