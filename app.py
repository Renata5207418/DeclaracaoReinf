import os
from datetime import datetime
import random
import string
import re
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from bson.objectid import ObjectId

from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, session, jsonify, make_response
)
from flask_pymongo import PyMongo
from flask_login import (
    LoginManager, UserMixin, login_user,
    login_required, logout_user, current_user
)
from flask_mail import Mail, Message
from werkzeug.security import generate_password_hash, check_password_hash
from itsdangerous import URLSafeTimedSerializer
from config import Config

app = Flask(__name__)
app.config.from_object(Config)

s = URLSafeTimedSerializer(app.config['SECRET_KEY'])

mongo = PyMongo(app)
mail = Mail(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'
login_manager.login_message = "Sessão expirada. Por favor, autentique-se novamente."
login_manager.login_message_category = "error"


def validate_cpf(cpf):
    cpf = re.sub(r'[^0-9]', '', cpf)
    if len(cpf) != 11: return False
    if cpf == cpf[0] * 11: return False
    soma = sum(int(cpf[i]) * (10 - i) for i in range(9))
    resto = (soma * 10) % 11
    if resto == 10 or resto == 11: resto = 0
    if resto != int(cpf[9]): return False
    soma = sum(int(cpf[i]) * (11 - i) for i in range(10))
    resto = (soma * 10) % 11
    if resto == 10 or resto == 11: resto = 0
    if resto != int(cpf[10]): return False
    return True


def validate_cnpj(cnpj):
    cnpj = re.sub(r'[^0-9]', '', cnpj)
    if len(cnpj) != 14: return False
    if cnpj == cnpj[0] * 14: return False
    tamanho = 12
    numeros = cnpj[:tamanho]
    digitos = cnpj[tamanho:]
    soma = 0
    pos = tamanho - 7
    for i in range(tamanho, 0, -1):
        soma += int(numeros[tamanho - i]) * pos
        pos -= 1
        if pos < 2: pos = 9
    resultado = 0 if soma % 11 < 2 else 11 - (soma % 11)
    if resultado != int(digitos[0]): return False
    tamanho = 13
    numeros = cnpj[:tamanho]
    soma = 0
    pos = tamanho - 7
    for i in range(tamanho, 0, -1):
        soma += int(numeros[tamanho - i]) * pos
        pos -= 1
        if pos < 2: pos = 9
    resultado = 0 if soma % 11 < 2 else 11 - (soma % 11)
    if resultado != int(digitos[1]): return False
    return True


def get_email_template(preheader, title, content_html, button_text=None, button_link=None):
    html = f"""
    <!DOCTYPE html>
    <html lang="pt-br">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
            body {{ margin: 0; padding: 0; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; background-color: #F3F4F6; color: #413D3A; }}
            .email-wrapper {{ width: 100%; background-color: #F3F4F6; padding: 40px 0; }}
            .email-container {{ max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 16px; overflow: hidden; box-shadow: 0 10px 30px rgba(0,0,0,0.08); }}
            .header {{ background-color: #ffffff; padding: 35px 40px; text-align: center; border-bottom: 5px solid #FBBA00; }}
            .logo-img {{ max-height: 50px; width: auto; display: block; margin: 0 auto; }}
            .content {{ padding: 45px 40px; background-color: #ffffff; }}
            .email-title {{ color: #413D3A; font-size: 24px; font-weight: 800; margin-top: 0; margin-bottom: 25px; letter-spacing: -0.5px; line-height: 1.3; }}
            .email-text {{ font-size: 16px; line-height: 1.6; color: #555555; margin-bottom: 25px; }}
            .btn-container {{ text-align: center; margin: 40px 0 20px 0; }}
            .btn {{ background-color: #FBBA00; color: #413D3A; padding: 18px 45px; text-decoration: none; font-weight: 800; border-radius: 10px; font-size: 14px; text-transform: uppercase; letter-spacing: 1px; display: inline-block; box-shadow: 0 4px 15px rgba(251, 186, 0, 0.4); }}
            .footer {{ background-color: #f9fafb; padding: 30px 40px; text-align: center; border-top: 1px solid #eeeeee; font-size: 12px; color: #999; }}
            .highlight-box {{ background-color: #fffbf0; border: 1px solid #ffeeba; border-left: 4px solid #FBBA00; padding: 25px; border-radius: 8px; margin: 30px 0; text-align: center; }}
            .token-code {{ font-family: 'Courier New', monospace; font-size: 36px; font-weight: 700; color: #413D3A; letter-spacing: 8px; display: block; margin: 0; }}
            .receipt-table {{ width: 100%; margin: 25px 0; border: 1px solid #eeeeee; border-radius: 8px; overflow: hidden; border-collapse: collapse; }}
            .receipt-row td {{ padding: 15px; border-bottom: 1px solid #eeeeee; font-size: 14px; color: #413D3A; }}
            .receipt-row:nth-child(even) {{ background-color: #fcfcfc; }}
            .receipt-value {{ font-weight: 700; text-align: right; white-space: nowrap; }}
        </style>
    </head>
    <body>
        <div class="email-wrapper">
            <div style="display:none;font-size:1px;color:#333;line-height:1px;max-height:0px;max-width:0px;opacity:0;overflow:hidden;">{preheader}</div>
            <div class="email-container">
                <div class="header">
                    <img src="cid:logo" alt="Scryta" class="logo-img">
                </div>
                <div class="content">
                    <h1 class="email-title">{title}</h1>
                    <div class="email-text">{content_html}</div>
                    {f'<div class="btn-container"><a href="{button_link}" class="btn">{button_text}</a></div>' if button_link else ''}
                </div>
                <div class="footer">
                    <p>&copy; {datetime.now().year} Scryta Contabilidade Digital.</p>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    return html


def send_email_with_logo(msg):
    try:
        with app.open_resource("static/img/logo.png") as fp:
            msg.attach("logo.png", "image/png", fp.read(),
                       headers={'Content-ID': '<logo>', 'Content-Disposition': 'inline'})
    except Exception:
        pass
    mail.send(msg)


def log_action(user_id, user_email, action, details=None):
    mongo.db.logs.insert_one({
        "user_id": ObjectId(user_id) if user_id else None,
        "email": user_email,
        "action": action,
        "details": details,
        "ip": request.remote_addr,
        "timestamp": datetime.utcnow()
    })


class User(UserMixin):
    def __init__(self, user_data):
        self.id = str(user_data['_id'])
        self.email = user_data['email']
        self.name = user_data['name']
        self.cpf = user_data.get('cpf')
        self.is_admin = user_data.get('is_admin', False)
        self.term_accepted_at = user_data.get('term_accepted_at')


@login_manager.user_loader
def load_user(user_id):
    try:
        user_data = mongo.db.users.find_one({"_id": ObjectId(user_id)})
        return User(user_data) if user_data else None
    except:
        return None


def is_password_strong(password):
    if len(password) < 8: return False
    if not re.search(r"[A-Z]", password): return False
    if not re.search(r"\d", password): return False
    return True


def generate_token():
    return ''.join(random.choices(string.digits, k=6))


@app.route('/')
def index():
    if current_user.is_authenticated: return redirect(url_for('dashboard'))
    return redirect(url_for('login'))


@app.route('/signup', methods=['GET', 'POST'])
def signup():
    if request.method == 'POST':
        email = request.form.get('email')
        if mongo.db.users.find_one({"email": email}):
            flash('Este e-mail corporativo já consta em nossa base.', 'error')
            return redirect(url_for('login'))

        token = s.dumps(email, salt='email-confirm')
        link = url_for('register_complete', token=token, _external=True)

        try:
            content = "<p>Recebemos uma solicitação para integrar sua empresa ao ecossistema Scryta.</p>"
            msg = Message('Convite Empresarial - Scryta', recipients=[email])
            msg.html = get_email_template("Ative sua conta.", "Bem-vindo", content, "ATIVAR CONTA", link)
            send_email_with_logo(msg)
            flash(f'Link enviado para {email}.', 'success')
        except Exception as e:
            flash(f'Erro SMTP: {str(e)}', 'error')

        return redirect(url_for('login'))
    return render_template('signup.html')


@app.route('/register/<token>', methods=['GET', 'POST'])
def register_complete(token):
    try:
        email = s.loads(token, salt='email-confirm', max_age=3600)
    except:
        flash('Link inválido ou expirado.', 'error')
        return redirect(url_for('signup'))

    if request.method == 'POST':
        name = request.form.get('name')
        cpf = request.form.get('cpf')
        password = request.form.get('password')
        confirm_password = request.form.get('confirm_password')
        terms = request.form.get('terms')

        if not terms:
            flash('Aceite os termos.', 'error')
            return redirect(request.url)
        if not validate_cpf(cpf):
            flash('CPF inválido.', 'error')
            return redirect(request.url)
        if password != confirm_password:
            flash('Senhas não conferem.', 'error')
            return redirect(request.url)
        if not is_password_strong(password):
            flash('Senha fraca.', 'error')
            return redirect(request.url)
        if mongo.db.users.find_one({"cpf": cpf}):
            flash('CPF já cadastrado.', 'error')
            return redirect(request.url)

        hashed = generate_password_hash(password)
        is_admin = True if mongo.db.users.count_documents({}) == 0 else False
        user_id = mongo.db.users.insert_one({
            "name": name, "cpf": cpf, "email": email, "password": hashed,
            "is_admin": is_admin, "created_at": datetime.utcnow(),
            "term_accepted_at": datetime.utcnow(), "status": "active"
        }).inserted_id

        log_action(user_id, email, "REGISTER", f"CPF: {cpf}")
        flash('Cadastro concluído!', 'success')
        return redirect(url_for('login'))
    return render_template('register.html', email=email, token=token)


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form.get('email')
        password = request.form.get('password')
        user_data = mongo.db.users.find_one({"email": email})

        if user_data and check_password_hash(user_data['password'], password):
            user = User(user_data)
            login_user(user)
            session.pop('batch_items', None)
            return redirect(url_for('dashboard'))
        else:
            flash('Credenciais inválidas.', 'error')
    return render_template('login.html')


@app.route('/forgot_password', methods=['GET', 'POST'])
def forgot_password():
    if request.method == 'POST':
        email = request.form.get('email')
        user = mongo.db.users.find_one({"email": email})
        if user:
            token = s.dumps(email, salt='password-reset')
            link = url_for('reset_password', token=token, _external=True)
            try:
                content = "<p>Você solicitou a redefinição de sua senha de acesso.</p>"
                msg = Message('Redefinição de Senha - Scryta', recipients=[email])
                msg.html = get_email_template("Redefinir Senha", "Recuperação", content, "CRIAR NOVA SENHA", link)
                send_email_with_logo(msg)
            except Exception as e:
                pass
        flash('Se o e-mail existir em nossa base, um link de recuperação foi enviado.', 'success')
        return redirect(url_for('login'))
    return render_template('forgot_password.html')


@app.route('/reset_password/<token>', methods=['GET', 'POST'])
def reset_password(token):
    try:
        email = s.loads(token, salt='password-reset', max_age=3600)
    except:
        flash('Link inválido ou expirado.', 'error')
        return redirect(url_for('login'))

    if request.method == 'POST':
        password = request.form.get('password')
        confirm_password = request.form.get('confirm_password')

        if password != confirm_password:
            flash('Senhas não conferem.', 'error')
            return redirect(request.url)
        if not is_password_strong(password):
            flash('Senha não atende aos critérios de segurança.', 'error')
            return redirect(request.url)

        hashed = generate_password_hash(password)
        mongo.db.users.update_one({"email": email}, {"$set": {"password": hashed}})
        flash('Senha atualizada com sucesso!', 'success')
        return redirect(url_for('login'))
    return render_template('reset_password.html', token=token)


@app.route('/logout')
@login_required
def logout():
    session.pop('batch_items', None)
    logout_user()
    return redirect(url_for('login'))


@app.route('/dashboard')
@login_required
def dashboard():
    if current_user.is_admin: return redirect(url_for('admin_panel'))
    
    user_companies = list(mongo.db.companies.find({"user_id": ObjectId(current_user.id)}))
    
    # Busca o histórico do usuário, EXCLUINDO os cancelados da visão do cliente
    user_history = list(mongo.db.user_financials.find({
        "user_id": ObjectId(current_user.id),
        "status": {"$ne": "desconsiderado"}
    }).sort("submitted_at", -1))
    
    response = make_response(render_template('dashboard.html', 
                                           name=current_user.name, 
                                           companies=user_companies,
                                           history=user_history))
    
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    
    return response


@app.route('/add_company', methods=['POST'])
@login_required
def add_company():
    company_cnpj = request.form.get('company_cnpj')

    if not company_cnpj:
        flash('CNPJ obrigatório.', 'error')
        return redirect(url_for('dashboard'))
    if not validate_cnpj(company_cnpj):
        flash('CNPJ inválido.', 'error')
        return redirect(url_for('dashboard'))

    exists = mongo.db.companies.find_one({"user_id": ObjectId(current_user.id), "cnpj": company_cnpj})
    if exists:
        flash('Empresa já cadastrada.', 'error')
    else:
        mongo.db.companies.insert_one({
            "user_id": ObjectId(current_user.id), "name": company_cnpj,
            "cnpj": company_cnpj, "created_at": datetime.utcnow()
        })
        flash('Empresa adicionada!', 'success')
    return redirect(url_for('dashboard'))


@app.route('/request_token', methods=['POST'])
@login_required
def request_token():
    data = request.json
    valor_str = data.get('valor', '0')
    sem_movimento = data.get('sem_movimento', False)
    data_retirada = data.get('data')
    company_id = data.get('company_id')
    confirmed_tax = data.get('confirmed_tax', False)
    action = data.get('action', 'finish')

    if sem_movimento:
        valor_atual = 0.0
        valor_str = 'R$ 0,00'
    else:
        try:
            valor_clean = valor_str.replace('R$', '').replace('.', '').replace(',', '.').strip()
            valor_atual = float(valor_clean)
        except:
            return jsonify({'status': 'error', 'message': 'Valor inválido.'}), 400

    if not data_retirada or not company_id:
        return jsonify({'status': 'error', 'message': 'Preencha todos os campos obrigatórios.'}), 400
    if not sem_movimento and not valor_atual:
        return jsonify({'status': 'error', 'message': 'Preencha o valor ou marque sem movimento.'}), 400

    company = mongo.db.companies.find_one({"_id": ObjectId(company_id), "user_id": ObjectId(current_user.id)})
    if not company:
        return jsonify({'status': 'error', 'message': 'Empresa inválida.'}), 400

    batch_items = session.get('batch_items', [])
    batch_total = sum(item['valor_numerico'] for item in batch_items)

    dt_obj = datetime.strptime(data_retirada, '%Y-%m-%d')
    start_date = datetime(dt_obj.year, dt_obj.month, 1)
    if dt_obj.month == 12:
        end_date = datetime(dt_obj.year + 1, 1, 1)
    else:
        end_date = datetime(dt_obj.year, dt_obj.month + 1, 1)

    # Cálculo dos 50k agora IGNORA os lançamentos desconsiderados pelo admin!
    pipeline = [
        {
            "$match": {
                "user_id": ObjectId(current_user.id),
                "status": {"$ne": "desconsiderado"},
                "data_retirada": {
                    "$gte": start_date.strftime('%Y-%m-%d'),
                    "$lt": end_date.strftime('%Y-%m-%d')
                }
            }
        },
        {"$group": {"_id": None, "total": {"$sum": "$valor_numerico"}}}
    ]
    resultado_agregacao = list(mongo.db.user_financials.aggregate(pipeline))
    historico_total = resultado_agregacao[0]['total'] if resultado_agregacao else 0.0

    total_projetado = historico_total + batch_total + valor_atual
    tax_details = None

    if total_projetado > 50000:
        if not confirmed_tax:
            base_calculo = total_projetado / 0.9
            imposto = base_calculo * 0.10
            liquido = total_projetado

            return jsonify({
                'status': 'warning_tax',
                'message': 'Limite de isenção excedido.',
                'calculations': {
                    'total_acumulado': f"R$ {total_projetado:,.2f}",
                    'base_calculo': f"R$ {base_calculo:,.2f}",
                    'imposto': f"R$ {imposto:,.2f}",
                    'liquido': f"R$ {liquido:,.2f}",
                    'aviso': 'O valor total de retiradas no mês excede R$ 50.000,00. O cálculo será aplicado sobre o TOTAL acumulado.'
                }
            })
        else:
            tax_details = {
                "base_calculo": total_projetado / 0.9,
                "imposto": (total_projetado / 0.9) * 0.10,
                "liquido_final": total_projetado,
                "total_acumulado_mes": total_projetado
            }

    current_item = {
        'valor_formatado': valor_str,
        'valor_numerico': valor_atual,
        'data_retirada': data_retirada,
        'company_id': str(company['_id']),
        'company_name': company['name'],
        'company_cnpj': company['cnpj'],
        'timestamp': datetime.utcnow().isoformat()
    }

    if action == 'add':
        batch_items.append(current_item)
        session['batch_items'] = batch_items
        return jsonify({'status': 'added', 'message': 'Empresa adicionada à lista. Insira a próxima.'})

    elif action == 'finish':
        batch_items.append(current_item)

        session['final_submission'] = {
            'items': batch_items,
            'tax_details': tax_details
        }

        token = generate_token()
        session['auth_token'] = token

        try:
            content = f"""
            <div class="highlight-box">
                <span style="font-size:12px; font-weight:bold; color:#999; text-transform:uppercase;">Código de Assinatura</span>
                <span class="token-code">{token}</span>
            </div>
            <p style="text-align:center;">Este código valida <strong>{len(batch_items)}</strong> declaração(ões) pendente(s).</p>
            """
            msg = Message("Token Scryta", recipients=[current_user.email])
            msg.html = get_email_template("Token de segurança", "Assinatura em Lote", content)
            send_email_with_logo(msg)
            return jsonify({'status': 'token_sent'})
        except Exception as e:
            return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/submit_withdrawal', methods=['POST'])
@login_required
def submit_withdrawal():
    data = request.json
    code = data.get('code')

    if code == session.get('auth_token'):
        sub_data = session.get('final_submission')
        if not sub_data: return jsonify({'status': 'error', 'message': 'Sessão expirada.'}), 400

        items = sub_data['items']
        tax_details = sub_data.get('tax_details')

        batch_id_db = generate_token()

        rows_html = ""

        for item in items:
            doc = {
                "user_id": ObjectId(current_user.id),
                "user_name": current_user.name,
                "user_cpf": current_user.cpf,
                "user_email": current_user.email,
                "company_name": item['company_name'],
                "company_cnpj": item['company_cnpj'],
                "valor": item['valor_formatado'],
                "valor_numerico": item['valor_numerico'],
                "data_retirada": item['data_retirada'],
                "submitted_at": datetime.utcnow(),
                "ip_address": request.remote_addr,
                "validation_token_used": True,
                "batch_id": batch_id_db,
                "fiscal_data_ref": tax_details,
                "status": "ativo"
            }
            mongo.db.user_financials.insert_one(doc)

            date_ptbr = datetime.strptime(item['data_retirada'], '%Y-%m-%d').strftime('%d/%m/%Y')
            rows_html += f"""
            <tr class="receipt-row">
                <td style="text-align:left;">{item['company_cnpj']}</td>
                <td>{date_ptbr}</td>
                <td class="receipt-value">{item['valor_formatado']}</td>
            </tr>
            """

        try:
            extra_html = ""
            if tax_details:
                extra_html = f"""
                <br>
                <div style="background-color:#fffbf0; border:1px solid #ffeeba; border-radius:8px; padding:15px; margin-top:20px;">
                    <h3 style="margin:0 0 10px 0; font-size:14px; text-transform:uppercase; color:#FBBA00;">Cálculo Fiscal Consolidado (Teto 50k)</h3>
                    <table style="width:100%; font-size:13px;">
                        <tr><td style="padding:5px 0;">Total Acumulado</td><td style="text-align:right; font-weight:bold;">R$ {tax_details['total_acumulado_mes']:,.2f}</td></tr>
                        <tr><td style="padding:5px 0;">Base (Gross-up)</td><td style="text-align:right; font-weight:bold;">R$ {tax_details['base_calculo']:,.2f}</td></tr>
                        <tr><td style="padding:5px 0;">IRRF (10%)</td><td style="text-align:right; font-weight:bold; color:red;">- R$ {tax_details['imposto']:,.2f}</td></tr>
                        <tr style="border-top:1px solid #ddd;"><td style="padding:8px 0; font-weight:bold;">Líquido Final</td><td style="text-align:right; font-weight:bold; color:green;">R$ {tax_details['liquido_final']:,.2f}</td></tr>
                    </table>
                </div>
                """

            receipt_html = f"""
            <p>Confirmamos o processamento de <strong>{len(items)}</strong> declaração(ões).</p>
            <table class="receipt-table">
                <thead><tr><th style="text-align:left; padding:10px;">Empresa</th><th>Data</th><th style="text-align:right;">Valor</th></tr></thead>
                <tbody>{rows_html}</tbody>
            </table>
            {extra_html}
            <p style="font-size:11px; color:#999; margin-top:20px; text-align:center;">ID do Lote: {batch_id_db}</p>
            """
            msg = Message("Comprovante - Scryta", recipients=[current_user.email])
            msg.html = get_email_template("Envio confirmado.", "Recibo Oficial", receipt_html)
            send_email_with_logo(msg)
        except Exception as e:
            print(f"Erro email: {e}")

        session.pop('auth_token', None)
        session.pop('final_submission', None)
        session.pop('batch_items', None)

        return jsonify({'status': 'success'})

    return jsonify({'status': 'error', 'message': 'Token inválido.'}), 400


# ==========================================
# ROTAS DO ADMIN (NOVO FLUXO DE SOFT DELETE)
# ==========================================

@app.route('/admin')
@login_required
def admin_panel():
    if not current_user.is_admin: return redirect(url_for('dashboard'))

    tab = request.args.get('tab', 'envios')
    all_users = list(mongo.db.users.find().sort("name", 1))

    if tab == 'envios':
        pipeline = [
            {
                "$match": { "status": {"$ne": "desconsiderado"} } # Ignora os cancelados
            },
            {
                "$addFields": {
                    "mes_ref": {"$substr": ["$data_retirada", 0, 7]}
                }
            },
            {
                "$sort": {"submitted_at": -1}
            },
            {
                "$group": {
                    "_id": {
                        "user_id": "$user_id",
                        "mes_ref": "$mes_ref"
                    },
                    "user_name": {"$first": "$user_name"},
                    "user_cpf": {"$first": "$user_cpf"},
                    "user_email": {"$first": "$user_email"},
                    "total_declarado": {"$sum": "$valor_numerico"},
                    "qtd_empresas": {"$sum": 1},
                    "detalhes": {
                        "$push": {
                            "id": {"$toString": "$_id"},
                            "empresa": "$company_name",
                            "cnpj": "$company_cnpj",
                            "valor": "$valor",
                            "data": "$data_retirada"
                        }
                    }
                }
            },
            {
                "$sort": {"_id.mes_ref": -1}
            }
        ]

        grouped_submissions = list(mongo.db.user_financials.aggregate(pipeline))

        for item in grouped_submissions:
            if item['total_declarado'] > 50000:
                base = item['total_declarado'] / 0.9
                imposto = base * 0.10
                liquido = item['total_declarado']
                item['calculo_dinamico'] = {
                    "base": base,
                    "imposto": imposto,
                    "liquido": liquido
                }
            else:
                item['calculo_dinamico'] = None

        return render_template('admin.html', grouped_data=grouped_submissions, active_tab='envios', all_users=all_users)

    elif tab == 'cancelados':
        # Busca apenas os lançamentos que o admin desconsiderou
        cancelados = list(mongo.db.user_financials.find({"status": "desconsiderado"}).sort("cancelled_at", -1))
        return render_template('admin.html', cancelados=cancelados, active_tab='cancelados', all_users=all_users)

    else:
        return render_template('admin.html', users=all_users, active_tab='users')


@app.route('/admin/record/add', methods=['POST'])
@login_required
def admin_add_record():
    if not current_user.is_admin: return redirect(url_for('dashboard'))
    
    user_id = request.form.get('user_id')
    company_cnpj = request.form.get('company_cnpj')
    valor_str = request.form.get('valor')
    data_retirada = request.form.get('data_retirada')

    target_user = mongo.db.users.find_one({"_id": ObjectId(user_id)})
    if not target_user:
        flash('Usuário não encontrado.', 'error')
        return redirect(url_for('admin_panel', tab='envios'))

    try:
        valor_clean = valor_str.replace('R$', '').replace('.', '').replace(',', '.').strip()
        valor_num = float(valor_clean)
    except:
        flash('Valor inválido.', 'error')
        return redirect(url_for('admin_panel', tab='envios'))

    doc = {
        "user_id": target_user['_id'],
        "user_name": target_user['name'],
        "user_cpf": target_user['cpf'],
        "user_email": target_user['email'],
        "company_name": "Lançamento Manual (Admin)",
        "company_cnpj": company_cnpj,
        "valor": valor_str,
        "valor_numerico": valor_num,
        "data_retirada": data_retirada,
        "submitted_at": datetime.utcnow(),
        "ip_address": request.remote_addr,
        "validation_token_used": False,
        "batch_id": "ADMIN_MANUAL",
        "fiscal_data_ref": None,
        "status": "ativo"
    }
    mongo.db.user_financials.insert_one(doc)
    flash('Lançamento manual inserido com sucesso.', 'success')
    return redirect(url_for('admin_panel', tab='envios'))


@app.route('/admin/record/edit', methods=['POST'])
@login_required
def admin_edit_record():
    if not current_user.is_admin: return redirect(url_for('dashboard'))
    
    record_id = request.form.get('record_id')
    novo_valor_str = request.form.get('valor')
    nova_data = request.form.get('data_retirada')

    try:
        valor_clean = novo_valor_str.replace('R$', '').replace('.', '').replace(',', '.').strip()
        valor_num = float(valor_clean)
    except:
        flash('Valor inválido.', 'error')
        return redirect(url_for('admin_panel', tab='envios'))

    mongo.db.user_financials.update_one(
        {"_id": ObjectId(record_id)},
        {"$set": {
            "valor": novo_valor_str, 
            "valor_numerico": valor_num, 
            "data_retirada": nova_data
        }}
    )
    flash('Registro atualizado com sucesso.', 'success')
    return redirect(url_for('admin_panel', tab='envios'))


@app.route('/admin/record/cancel/<record_id>', methods=['POST'])
@login_required
def admin_cancel_record(record_id):
    if not current_user.is_admin: return redirect(url_for('dashboard'))
    
    # Em vez de deletar, atualiza o status (Soft Delete)
    mongo.db.user_financials.update_one(
        {"_id": ObjectId(record_id)},
        {"$set": {
            "status": "desconsiderado",
            "cancelled_at": datetime.utcnow()
        }}
    )
    flash('Lançamento movido para desconsiderados.', 'success')
    return redirect(url_for('admin_panel', tab='envios'))


@app.route('/admin/record/restore/<record_id>', methods=['POST'])
@login_required
def admin_restore_record(record_id):
    if not current_user.is_admin: return redirect(url_for('dashboard'))
    
    # Remove a flag de desconsiderado, ativando novamente
    mongo.db.user_financials.update_one(
        {"_id": ObjectId(record_id)},
        {"$set": {"status": "ativo"}, "$unset": {"cancelled_at": ""}}
    )
    flash('Lançamento restaurado com sucesso. Ele voltou para os cálculos.', 'success')
    return redirect(url_for('admin_panel', tab='cancelados'))


@app.route('/admin/user/toggle/<user_id>', methods=['POST'])
@login_required
def admin_toggle_user(user_id):
    if not current_user.is_admin: return redirect(url_for('dashboard'))
    
    if user_id == str(current_user.id):
        flash('Você não pode alterar seu próprio nível de acesso.', 'error')
        return redirect(url_for('admin_panel', tab='users'))

    target_user = mongo.db.users.find_one({"_id": ObjectId(user_id)})
    if target_user:
        current_status = target_user.get('is_admin', False)
        mongo.db.users.update_one({"_id": ObjectId(user_id)}, {"$set": {"is_admin": not current_status}})
        msg = f"Usuário {target_user['name']} agora é Admin." if not current_status else f"Acesso Admin removido de {target_user['name']}."
        flash(msg, 'success')
        
    return redirect(url_for('admin_panel', tab='users'))


@app.route('/admin/term_proof/<user_id>')
@login_required
def term_proof(user_id):
    if not current_user.is_admin: return redirect(url_for('dashboard'))

    user = mongo.db.users.find_one({"_id": ObjectId(user_id)})
    if not user: return "Usuário não encontrado", 404

    # Relatório em PDF também ignora os desconsiderados
    financials = list(mongo.db.user_financials.find({
        "user_id": ObjectId(user_id),
        "status": {"$ne": "desconsiderado"}
    }).sort("data_retirada", -1))

    return render_template('print_proof.html', user=user, financials=financials, now=datetime.utcnow())


@app.route('/admin/export')
@login_required
def export_excel():
    if not current_user.is_admin: return redirect(url_for('dashboard'))

    wb = openpyxl.Workbook()

    ws_resumo = wb.active
    ws_resumo.title = "Total Lotes"

    ws_detalhes = wb.create_sheet(title="Registros Individuais")

    header_fill = PatternFill(start_color="413D3A", end_color="413D3A", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    border_style = Border(
        left=Side(style='thin', color='E5E7EB'), right=Side(style='thin', color='E5E7EB'),
        top=Side(style='thin', color='E5E7EB'), bottom=Side(style='thin', color='E5E7EB')
    )
    center_align = Alignment(horizontal="center", vertical="center")

    headers_resumo = ['Mês Referência', 'Nome do Usuário', 'CPF', 'Total Acumulado Lote', 'Base Cálculo',
                      'Imposto Retido', 'Líquido Recebido']
    ws_resumo.append(headers_resumo)
    for col_num, cell in enumerate(ws_resumo[1], 1):
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align

    pipeline = [
        {"$match": {"status": {"$ne": "desconsiderado"}}}, # Planilha Excel ignora os cancelados
        {"$addFields": {"mes_ref": {"$substr": ["$data_retirada", 0, 7]}}},
        {"$group": {
            "_id": {"user_id": "$user_id", "mes_ref": "$mes_ref"},
            "user_name": {"$first": "$user_name"},
            "user_cpf": {"$first": "$user_cpf"},
            "total": {"$sum": "$valor_numerico"}
        }},
        {"$sort": {"_id.mes_ref": -1}}
    ]
    resumo_data = list(mongo.db.user_financials.aggregate(pipeline))

    for r in resumo_data:
        total = r['total']
        if total > 50000:
            base = total / 0.9
            imp = base * 0.10
            liq = total
        else:
            base = 0
            imp = 0
            liq = total

        row = [r['_id']['mes_ref'], r['user_name'], r['user_cpf'], total, base, imp, liq]
        ws_resumo.append(row)

    for row in ws_resumo.iter_rows(min_row=2, max_col=7):
        for cell in row:
            cell.border = border_style
        for idx in [3, 4, 5, 6]:
            row[idx].number_format = '"R$" #,##0.00'

    for col in ws_resumo.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        ws_resumo.column_dimensions[column].width = max_length + 6

    headers_detalhes = ['ID Lote', 'CNPJ Empresa', 'Nome Usuário', 'CPF', 'Data Lançamento', 'Valor Lançamento']
    ws_detalhes.append(headers_detalhes)
    for col_num, cell in enumerate(ws_detalhes[1], 1):
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align

    # Planilha Excel (Aba Detalhes) também ignora cancelados
    financial_records = list(mongo.db.user_financials.find({"status": {"$ne": "desconsiderado"}}).sort([("data_retirada", -1), ("company_cnpj", 1)]))
    for s in financial_records:
        row = [
            s.get('batch_id', 'N/A'),
            s.get('company_cnpj', ''),
            s.get('user_name', ''),
            s.get('user_cpf', ''),
            s.get('data_retirada', ''),
            s.get('valor_numerico', 0)
        ]
        ws_detalhes.append(row)

    for row in ws_detalhes.iter_rows(min_row=2, max_col=6):
        for cell in row:
            cell.border = border_style
        row[5].number_format = '"R$" #,##0.00'

    for col in ws_detalhes.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        ws_detalhes.column_dimensions[column].width = max_length + 6

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    resp = make_response(output.getvalue())
    resp.headers["Content-Disposition"] = "attachment; filename=relatorio_completo_scryta.xlsx"
    resp.headers["Content-type"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return resp


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5894, debug=True)