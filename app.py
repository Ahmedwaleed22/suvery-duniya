from flask import Flask, render_template, request, jsonify, send_from_directory, redirect, url_for
from flask_cors import CORS
from flask_restful import abort
from werkzeug.utils import secure_filename
from flask_mysqldb import MySQL
from functools import wraps
from subprocess import Popen, PIPE
import datetime
import stripe
import os
import jwt

stripe_keys = {
  'secret_key': 'sk_test_51KL5fZSIeo8cbA9acaGiriXApreMJoliz1YnTX9SdweJMliwxVUeMDoU8zdb8nLFFVMmcUizp1S2BSlx1OVwhJ5j00i59tU6US',
  'publishable_key': 'pk_test_51KL5fZSIeo8cbA9aFNsoJYy35jp4ePQtnrJOSUwc34oMGcc0KL88oW1KB9jsdm0ZoShZwOh1zLba7yI6c6NxDIc200kIlLxI9q'
}

stripe.api_key = stripe_keys['secret_key']

app = Flask(__name__, static_url_path='/static')
SECRET_KEY = os.environ.get('SECRET_KEY') or '5j$@n324h@&$98redjrsdg43jh4*32$34&@5u4'

app.config['SECRET_KEY'] = SECRET_KEY
app.config['MYSQL_HOST'] = '127.0.0.1'
app.config['MYSQL_USER'] = 'root'
app.config['MYSQL_PASSWORD'] = ''
app.config['MYSQL_DB'] = 'tasksDashboard'

mysql = MySQL(app)
cors = CORS(app, resources={"*": {"origins": "*"}})

path = os.getcwd()
UPLOAD_FOLDER = os.path.join(path, 'models')
CSV_FOLDER = os.path.join(path, 'csv')
EXCEL_FOLDER = os.path.join(path, 'excel')
WORD_FOLDER = os.path.join(path, 'word')

if not os.path.isdir(UPLOAD_FOLDER):
  os.mkdir(UPLOAD_FOLDER)

if not os.path.isdir(CSV_FOLDER):
  os.mkdir(CSV_FOLDER)

if not os.path.isdir(EXCEL_FOLDER):
  os.mkdir(EXCEL_FOLDER)

if not os.path.isdir(WORD_FOLDER):
  os.mkdir(WORD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['EXCEL_FOLDER'] = EXCEL_FOLDER
app.config['WORD_FOLDER'] = WORD_FOLDER

class CursorByName():
  def __init__(self, cursor):
    self._cursor = cursor
  
  def __iter__(self):
    return self

  def __next__(self):
    row = self._cursor.__next__()
    return { description[0]: row[col] for col, description in enumerate(self._cursor.description) }


def token_required(f):
  @wraps(f)
  def decorated(*args, **kwargs):
    token = None
    if "Authorization" in request.headers:
      token = request.headers["Authorization"].split(" ")[1]
    if not token:
      return {
        "message": "Authentication Token is missing!",
        "data": None,
        "error": "Unauthorized"
      }, 401
    try:
      data = jwt.decode(token, app.config["SECRET_KEY"], algorithms=["HS256"])

      cursor = mysql.connection.cursor()
      cursor.execute("SELECT * FROM users WHERE id = %s", (data['public_id'],))
      current_user = cursor.fetchone()
      cursor.close()

      user = {
        "id": current_user[0],
        "username": current_user[1],
      }

      if current_user is None:
        return {
          "message": "Invalid Authentication token!",
          "data": None,
          "error": "Unauthorized"
        }, 401
    except Exception as e:
      return {
        "message": "Something went wrong",
        "data": None,
        "error": str(e)
      }, 500

    return f(user, *args, **kwargs)

  return decorated

@app.route("/static/<path:path>")
def static_dir(path):
  return send_from_directory("static", path)

@app.errorhandler(404)
def not_found(e):
  return render_template('index.html')


@app.route('/', methods=['GET'])
def index():
  return render_template('index.html')


@app.route('/api/login', methods=['POST'])
def login():
  data = request.get_json()
  username = data['username']
  password = data['password']
  cursor = mysql.connection.cursor()
  cursor.execute("SELECT * FROM users WHERE username = %s", (username,))
  user = cursor.fetchone()
  cursor.close()
  if user is None:
    return jsonify({
      "message": "User not found!",
      "data": None,
      "error": "Unauthorized"
    }), 401
  if user[2] != password:
    return jsonify({
      "message": "Invalid password!",
      "data": None,
      "error": "Unauthorized"
    }), 401
  token = jwt.encode({'public_id': user[0], 'exp' : datetime.datetime.utcnow() + datetime.timedelta(minutes=30)}, app.config['SECRET_KEY'])  

  return jsonify({
    "message": "Successfully logged in!",
    "data": {
      "token": token
    },
    "error": None
  }), 200


@app.route('/api/token/check', methods=['POST'])
@token_required
def check_token(user):
  return jsonify({
    "message": "Successfully logged in!",
    "data": {
      "token": user
    },
    "error": None
  }), 200


# @app.route('/api/survies', methods=['GET'])
# def get_survies():
#   cursor = mysql.connection.cursor()
#   cursor.execute("SELECT google_form_survies.*, customers.username as username FROM google_form_survies INNER JOIN customers ON customers.ID = google_form_survies.customer_id WHERE google_form_survies.status = 1 OR google_form_survies.status = 2")
#   survies = cursor.fetchall()
#   columns = [column[0] for column in cursor.description]
#   result = []
#   cursor.close()

#   for value in survies:
#     tmp = {}
#     for (index, column) in enumerate(value):
#         tmp[columns[index]] = column
#     result.append(tmp)

#   return jsonify({
#     "message": "Successfully fetched survies!",
#     "data": {
#       "survies": result
#     },
#     "error": None
#   }), 200


# @app.route('/api/survies/<int:id>', methods=['GET'])
# def get_survie(id):
#   cursor = mysql.connection.cursor()
#   cursor.execute("SELECT * FROM survies WHERE id = %s", (id,))
#   survie = cursor.fetchone()
#   columns = [column[0] for column in cursor.description]
#   result = {}
#   for (index, column) in enumerate(survie):
#     result[columns[index]] = column
#   cursor.close()
#   return jsonify({
#     "message": "Successfully fetched survie!",
#     "data": {
#       "survey": result
#     },
#     "error": None
#   }), 200


# @app.route('/api/survies/<int:id>', methods=['PUT'])
# @token_required
# def modify_survies(user, id):
#   data = request.get_json()
#   survey = data['survey']
#   cursor = mysql.connection.cursor()
#   cursor.execute("UPDATE survies SET survey = %s, status = %s WHERE id = %s", (survey, 4, id))
#   mysql.connection.commit()
#   cursor.close()
#   return jsonify({
#     "message": "Successfully modified survey!",
#     "data": {
#       "survey": survey
#     },
#     "error": None
#   }), 200


# @app.route('/api/survies/<int:id>/reject', methods=['POST'])
# @token_required
# def reject_survey(user, id):
#   cursor = mysql.connection.cursor()
#   cursor.execute("UPDATE google_form_survies SET status = %s WHERE id = %s", (3, id))
#   mysql.connection.commit()
#   cursor.close()
#   return jsonify({
#     "message": "Successfully rejected survey!",
#     "error": None
#   }), 200


# @app.route('/api/survies/<int:id>/approve', methods=['POST'])
# @token_required
# def approve_survey(user, id):
#   cursor = mysql.connection.cursor()
#   cursor.execute("UPDATE google_form_survies SET status = %s WHERE id = %s", (2, id))
#   mysql.connection.commit()
#   cursor.close()
#   return jsonify({
#     "message": "Successfully approved survey!",
#     "error": None
#   }), 200


# @app.route('/api/survies/<int:id>/comment', methods=['POST'])
# @token_required
# def comment_survey(user, id):
#   data = request.get_json()
#   comments = data['comments']
#   cursor = mysql.connection.cursor()
#   cursor.execute("UPDATE survies SET status = %s, comments = %s WHERE id = %s", (3, comments, id))
#   mysql.connection.commit()
#   cursor.close()
#   return jsonify({
#     "message": "Successfully approved survey!",
#     "error": None
#   }), 200


# # get comments by survey id
# @app.route('/api/survies/<int:id>/comments', methods=['GET'])
# @token_required
# def get_comments(user, id):
#   cursor = mysql.connection.cursor()
#   cursor.execute("SELECT comments FROM survies WHERE id = %s", (id,))
#   comments = cursor.fetchone()
#   cursor.close()
#   return jsonify({
#     "message": "Successfully fetched comments!",
#     "data": {
#       "comments": comments[0]
#     },
#     "error": None
#   }), 200


# # Insert into answers in database
# @app.route('/api/survies/<int:id>/answers', methods=['POST'])
# def insert_answers(id):
#   data = request.get_json()
#   answers = data['answers']
#   # cursor = mysql.connection.cursor()
#   # cursor.execute("INSERT INTO answers (survey_id, answers) VALUES (%s, %s)", (id, answers))
#   # mysql.connection.commit()
#   # cursor.close()
#   # return jsonify({
#   #   "message": "Successfully inserted answers!",
#   #   "data": {
#   #     "answers": answers
#   #   },
#   #   "error": None
#   # }), 200
#   return jsonify({
#     "message": "Successfully inserted answers!",
#     "data": {
#       "answers": answers
#     },
#     "error": None
#   }), 200


# @app.route('/api/customer/register', methods=['POST'])
# def customerRegister():
#   username = request.form.get('username')
#   email = request.form.get('email')
#   password = request.form.get('password')
#   option = request.form.get('option')
#   title = request.form.get('title')
#   targets = request.form.get('targets')
#   file_path = ""

#   if username is None or email is None or password is None or option is None or title is None or targets is None:
#     return abort(400)

#   if option == 'excel' or option == 'word':
#     file = request.files['file']

#     if file is None:
#       return abort(400)

#     if option == 'excel':
#       if file.filename.split('.')[1] in ['xlsx', 'xlsm', 'xlsb', 'xltx', 'csv']:
#         file_path = os.path.join(app.config['EXCEL_FOLDER'], secure_filename(file.filename))
#         file.save(file_path)

#       else:
#         return abort(400)

#     elif option == 'word':
#       if file.filename.split('.')[1] == 'docx':
#         file_path = os.path.join(app.config['WORD_FOLDER'], secure_filename(file.filename))
#         file.save(file_path)

#       else:
#         return abort(400)

#     output = Popen(["python3", f'option_models/option1and2_file_upload.py', file_path, title], stdout=PIPE)
#     response, err = output.communicate()
#     dataOutput = response.decode('utf-8').split('|')
#     google_form_id = dataOutput[2]
#     google_form_link = dataOutput[0]
#     questions_number = dataOutput[1]

#     cursor = mysql.connection.cursor()
#     cursor.execute("INSERT INTO customers (username, email, password) VALUES (%s, %s, %s)", (username, email, password))
#     mysql.connection.commit()
#     cursor.execute("INSERT INTO google_form_survies (google_form_id, google_form_link, target, questions_number, customer_id) VALUES (%s, %s, %s, %s, %s)", (google_form_id, google_form_link, targets, questions_number, cursor.lastrowid))
#     mysql.connection.commit()
#     cursor.close()

#   elif option == 'ui':
#     dat = request.form.get('dat')

#     output = Popen(["python3", f'option_models/option3_dynamic.py', dat, title], stdout=PIPE)
#     response, err = output.communicate()
#     dataOutput = response.decode('utf-8').split('|')
#     google_form_id = dataOutput[2]
#     google_form_link = dataOutput[0]
#     questions_number = dataOutput[1]

#     cursor = mysql.connection.cursor()
#     cursor.execute("INSERT INTO customers (username, email, password) VALUES (%s, %s, %s)", (username, email, password))
#     mysql.connection.commit()
#     cursor.execute("INSERT INTO google_form_survies (google_form_id, google_form_link, target, questions_number, customer_id) VALUES (%s, %s, %s, %s, %s)", (google_form_id, google_form_link, targets, questions_number, cursor.lastrowid))
#     mysql.connection.commit()
#     cursor.close()

#   return redirect('http://localhost:9090/')


# @app.route("/api/stripe/config")
# def get_publishable_key():
#   stripe_config = {"publicKey": stripe_keys["publishable_key"]}
#   return jsonify(stripe_config)


# @app.route('/api/checkout/<amount>/<user_id>', methods=['GET'])
# def checkoutApi(amount, user_id):
#   if int(amount) >= 1000:
#     product = stripe.Product.create(
#       name='Deposit Of ' + str(amount),
#     )

#     # checkout using stripe
#     price = stripe.Price.create(
#       currency='inr',
#       unit_amount=int(amount) * 100,
#       product=product.id
#     )

#     session = stripe.checkout.Session.create(
#       payment_method_types=['card'],
#       mode='payment',
#       line_items=[
#         {
#           'price': price.id,
#           'quantity': 1,
#         }
#       ],
#       payment_intent_data={
#         'metadata': {
#           'user_id': user_id,
#           'amount': amount
#         }
#       },
#       success_url='http://localhost:8000/checkout/success/{CHECKOUT_SESSION_ID}',
#       cancel_url=url_for('cancel', _external=True)
#     )

#     return jsonify({"sessionId": session["id"]})

#   return render_template('error.html')


# @app.route('/checkout/success/<session_id>', methods=['GET'])
# def success(session_id):
#   try:
#     session = stripe.checkout.Session.retrieve(session_id)
#     payment_intent = stripe.PaymentIntent.retrieve(session.payment_intent)
#     user_id = payment_intent.metadata.get('user_id')
#     amount = payment_intent.metadata.get('amount')

#     if session.payment_status == 'paid':
#       cursor = mysql.connection.cursor()
#       cursor.execute('SELECT balance FROM customers WHERE ID = %s', (user_id,))
#       balance = cursor.fetchone()[0]
#       cursor.close()

#       new_balance = balance + int(amount)
#       cursor = mysql.connection.cursor()
#       cursor.execute('UPDATE customers SET balance = %s WHERE ID = %s', (new_balance, user_id))
#       mysql.connection.commit()
#       cursor.close()

#       return redirect('http://localhost:9090/payment.php')

#     return render_template('error.html')
#   except stripe.error.InvalidRequestError:
#     return render_template('canceled.html')


# @app.route('/checkout/canceled', methods=['GET'])
# def cancel():
#   return render_template('canceled.html')


# @app.route('/api/pay/sendforreview/<survey_id>/<customer_id>', methods=['GET'])
# def sendSurveyForReview(survey_id, customer_id):
#   cursor = mysql.connection.cursor()
#   cursor.execute("SELECT ID, balance FROM customers WHERE ID = %s", (customer_id,))
#   user_data = cursor.fetchone()
#   balance = user_data[1]

#   cursor.execute("SELECT target, questions_number, status, customer_id FROM google_form_survies WHERE ID = %s", (survey_id,))
#   google_form_survey_data = cursor.fetchone()

#   if google_form_survey_data[3] != user_data[0]:
#     return jsonify({
#       "message": "User is not the owner of the survey",
#       "new_balance": balance
#     })

#   if google_form_survey_data[2] != 0:
#     return jsonify({
#       "message": "Survey not in pending status",
#       "new_balance": balance
#     })

#   survey_price = int(google_form_survey_data[0]) * int(google_form_survey_data[1]) * 10

#   if int(survey_price) > int(balance):
#     return jsonify({
#       "message": "Not enough balance",
#       "new_balance": balance
#     }), 406

#   # new_balance = balance - survey_price

#   # cursor.execute("UPDATE customers SET balance = %s WHERE ID = %s", (new_balance, customer_id,))
#   cursor.execute("UPDATE google_form_survies SET status = %s WHERE ID = %s", (1, survey_id))
#   mysql.connection.commit()


#   return jsonify({
#     "message": "Sent to admins for approval",
#     "new_balance": str(balance)
#   })


# @app.route('/api/charge/<customer_id>')
# def charge_user(customer_id):
#   cursor = mysql.connection.cursor()
#   cursor.execute("SELECT * FROM google_form_survies WHERE customer_id = %s AND status = 2", (customer_id,))
#   rows = cursor.fetchall()
#   total_price = 0

#   for row in rows:
#     if int(row[4]) > 10:
#       extra_questions = int(row[4]) - 10
#       questions = 10
#     else:
#       extra_questions = 0
#       questions = int(row[4])

#     output = Popen(["python3", f'option_models/create_dashboard.py', row[1]], stdout=PIPE)
#     response, err = output.communicate()
#     respondents_no = response.decode('utf-8')
    
#     price = ((questions * 5) + (extra_questions * 2)) * respondents_no

#     total_price += price

#   cursor.execute("SELECT * FROM customers WHERE ID = %s", (customer_id,))
#   user = cursor.fetchone()
#   amount_charged = user[5]

#   if int(total_price) > int(amount_charged):
#     amount_to_charge = int(total_price) - int(amount_charged)
#     new_balance = int(user[4]) - amount_to_charge
#     new_charged_amount = amount_to_charge + amount_charged

#     cursor.execute("UPDATE customers SET amount_charged = %s, balance = %s WHERE ID = %s", (new_charged_amount, new_balance, customer_id))
#     mysql.connection.commit()
#     cursor.close()

#     return jsonify({
#       "message": "",
#       "new_balance": str(new_balance)
#     })

#   new_balance = int(user[4])
#   return jsonify({
#       "message": "",
#       "new_balance": str(new_balance)
#     })


# Pages API Routes

# Pages create
@app.route('/api/pages', methods=['POST'])
@token_required
def create_page(user):
    page_name = request.json.get('name')
    page_meta_tags = request.json.get('meta_tags')
    page_content = request.json.get('content')
    page_title = request.json.get('title')
    page_meta_description = request.json.get('meta_description')
    category = request.json.get('category')

    cur = mysql.connection.cursor()
    cur.execute("INSERT INTO pages (name, meta_tags, title, meta_description, content, category) VALUES (%s, %s, %s, %s, %s, %s)",
                (page_name, page_meta_tags, page_title, page_meta_description, page_content, category))
    mysql.connection.commit()
    cur.close()

    return jsonify({
        "message": "Page created successfully"
    }), 201


# Pages list
@app.route('/api/pages')
@token_required
def all_pages(user):
    cursor = mysql.connection.cursor()
    cursor.execute("SELECT * FROM pages")
    columns = [column[0] for column in cursor.description]
    result = []

    for value in cursor.fetchall():
        tmp = {}
        for (index, column) in enumerate(value):
            tmp[columns[index]] = column
        result.append(tmp)

    cursor.close()
    return jsonify(result)


# View page
@app.route('/api/pages/<id>')
@token_required
def view_page(user, id):
    cursor = mysql.connection.cursor()
    cursor.execute("SELECT * FROM pages WHERE ID = %s", (id,))
    page = cursor.fetchone()
    cursor.close()

    return jsonify({
        "ID": page[0],
        "name": page[1],
        "meta_tags": page[2],
        "title": page[3],
        "description": page[4],
        "content": page[5],
        "category": page[6]
    })


# Update page
@app.route('/api/pages/<id>', methods=['PATCH'])
@token_required
def update_page(user, id):
    name = request.json.get("name")
    meta_tags = request.json.get("meta_tags")
    content = request.json.get("content")
    page_title = request.json.get('title')
    page_meta_description = request.json.get('meta_description')
    category = request.json.get('category')

    cursor = mysql.connection.cursor()
    cursor.execute("UPDATE pages SET name = %s, meta_tags = %s, title = %s, meta_description = %s, content = %s, category = %s WHERE ID = %s", (name, meta_tags, page_title, page_meta_description, content, category, id))
    mysql.connection.commit()
    cursor.close()

    return jsonify({
        "ID": id,
        "name": name,
        "meta_tags": meta_tags,
        "title": page_title,
        "description": page_meta_description,
        "content": content,
        "category": category
    })


# Delete page
@app.route('/api/pages/<id>', methods=['DELETE'])
@token_required
def delete_page(user, id):
    cursor = mysql.connection.cursor()
    cursor.execute("DELETE FROM pages WHERE ID = %s", (id,))
    mysql.connection.commit()
    cursor.close()

    return jsonify({
        "message": "Page Successfully Deleted"
    })


@app.route('/blog/<id>')
def show_post_page(id):
    cursor = mysql.connection.cursor()
    cursor.execute("SELECT * FROM pages WHERE ID = %s", (id,))
    page = cursor.fetchone()
    cursor.close()

    return render_template('show_post.html', page=page)


if __name__ == "__main__":
  app.run(debug=True, port=8000)