# ============================================================
# SISTEM ANTRIAN KLINIK - ANTRI CEPAT ASSIST.ID V3.0 (FINAL + FORMAT RUPIAH)
# ============================================================

import os
import json
import io
import qrcode
from datetime import datetime, timedelta
from functools import wraps

from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, session, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_socketio import SocketIO, emit
from apscheduler.schedulers.background import BackgroundScheduler
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook

app = Flask(__name__)
app.config["SECRET_KEY"] = "antri-cepat-assist-secret-key-2026"
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///klinik.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)
socketio = SocketIO(app, cors_allowed_origins="*", async_mode="threading")

# ═══════════════════════════════════════════════════════════════
# ✅ CUSTOM FILTER JINJA2 (Untuk format Rupiah yang aman)
# ═══════════════════════════════════════════════════════════════
@app.template_filter('rupiah')
def format_rupiah(angka):
    """Format angka menjadi format Rupiah (contoh: 150000 -> 150.000)"""
    try:
        angka_float = float(angka)
        return "{:,.0f}".format(angka_float).replace(",", ".")
    except (ValueError, TypeError):
        return "0"
# ═══════════════════════════════════════════════════════════════

# ── Model Database ───────────────────────────────────────────
class User(db.Model):
    __tablename__ = "users"
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    role = db.Column(db.String(20), default="admin")

class Doctor(db.Model):
    __tablename__ = "doctors"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    specialization = db.Column(db.String(100), nullable=False)
    schedule = db.Column(db.Text, nullable=False)
    quota = db.Column(db.Integer, default=30)
    consultation_fee = db.Column(db.Float, default=150000)
    is_active = db.Column(db.Boolean, default=True)
    appointments = db.relationship("Appointment", backref="doctor", lazy=True)

    def get_schedule_dict(self):
        try: return json.loads(self.schedule)
        except: return {}
        
class Patient(db.Model):
    __tablename__ = "patients"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    phone = db.Column(db.String(20), nullable=False)
    email = db.Column(db.String(120), default="")
    medical_record_number = db.Column(db.String(20), unique=True, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    appointments = db.relationship("Appointment", backref="patient", lazy=True)
    medical_records = db.relationship("MedicalRecord", backref="patient", lazy=True)
    notifications = db.relationship("Notification", backref="patient", lazy=True)

class Appointment(db.Model):
    __tablename__ = "appointments"
    id = db.Column(db.Integer, primary_key=True)
    patient_id = db.Column(db.Integer, db.ForeignKey("patients.id"), nullable=False)
    doctor_id = db.Column(db.Integer, db.ForeignKey("doctors.id"), nullable=False)
    appointment_date = db.Column(db.Date, nullable=False)
    queue_number = db.Column(db.Integer, nullable=False)
    status = db.Column(db.String(20), default="waiting")
    pharmacy_status = db.Column(db.String(20), default="waiting")
    qr_code = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    invoice = db.relationship("Invoice", backref="appointment", uselist=False, lazy=True)

class MedicalRecord(db.Model):
    __tablename__ = "medical_records"
    id = db.Column(db.Integer, primary_key=True)
    patient_id = db.Column(db.Integer, db.ForeignKey("patients.id"), nullable=False)
    appointment_id = db.Column(db.Integer, db.ForeignKey("appointments.id"), nullable=True)
    diagnosis = db.Column(db.Text, default="")
    treatment = db.Column(db.Text, default="")
    notes = db.Column(db.Text, default="")
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class Notification(db.Model):
    __tablename__ = "notifications"
    id = db.Column(db.Integer, primary_key=True)
    patient_id = db.Column(db.Integer, db.ForeignKey("patients.id"), nullable=False)
    message = db.Column(db.Text, nullable=False)
    is_read = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class Medicine(db.Model):
    __tablename__ = "medicines"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    category = db.Column(db.String(50), default="Umum")
    stock = db.Column(db.Integer, default=0)
    price = db.Column(db.Float, default=0)
    unit = db.Column(db.String(20), default="tablet")
    expiry_date = db.Column(db.Date, nullable=True)
    min_stock = db.Column(db.Integer, default=10)

class Invoice(db.Model):
    __tablename__ = "invoices"
    id = db.Column(db.Integer, primary_key=True)
    appointment_id = db.Column(db.Integer, db.ForeignKey("appointments.id"), unique=True, nullable=False)
    patient_id = db.Column(db.Integer, db.ForeignKey("patients.id"), nullable=False)
    consultation_fee = db.Column(db.Float, default=0)
    medicine_fee = db.Column(db.Float, default=0)
    total = db.Column(db.Float, default=0)
    status = db.Column(db.String(20), default="unpaid")
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    items = db.relationship("InvoiceItem", backref="invoice", lazy=True)

class InvoiceItem(db.Model):
    __tablename__ = "invoice_items"
    id = db.Column(db.Integer, primary_key=True)
    invoice_id = db.Column(db.Integer, db.ForeignKey("invoices.id"), nullable=False)
    description = db.Column(db.String(200), nullable=False)
    quantity = db.Column(db.Integer, default=1)
    price = db.Column(db.Float, default=0)
    total = db.Column(db.Float, default=0)

class Feedback(db.Model):
    __tablename__ = "feedbacks"
    id = db.Column(db.Integer, primary_key=True)
    appointment_id = db.Column(db.Integer, db.ForeignKey("appointments.id"), nullable=False)
    patient_id = db.Column(db.Integer, db.ForeignKey("patients.id"), nullable=False)
    rating = db.Column(db.Integer, nullable=False)
    comment = db.Column(db.Text, default="")
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    appointment = db.relationship("Appointment", backref="feedbacks", lazy=True)
    
# ── Helper Functions ─────────────────────────────────────────
def generate_mr_number():
    return f"MRN-{datetime.now().strftime('%Y%m')}-{Patient.query.count() + 1:04d}"

def get_next_queue(doctor_id, date):
    last = Appointment.query.filter_by(doctor_id=doctor_id, appointment_date=date).filter(Appointment.status != "cancelled").order_by(Appointment.queue_number.desc()).first()
    return (last.queue_number + 1) if last else 1

def generate_qr_code(appointment_id):
    qr_data = f"AULIA-KLINIK-CHECKIN-{appointment_id}"
    qr = qrcode.QRCode(version=1, box_size=10, border=5)
    qr.add_data(qr_data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    qr_dir = os.path.join(app.static_folder, 'qr_codes')
    os.makedirs(qr_dir, exist_ok=True)
    filename = f"qr_{appointment_id}.png"
    filepath = os.path.join(qr_dir, filename)
    img.save(filepath)
    return f"qr_codes/{filename}"

def create_notification(patient_id, message):
    notif = Notification(patient_id=patient_id, message=message)
    db.session.add(notif)
    db.session.commit()
    socketio.emit("new_notification", {"patient_id": patient_id, "message": message})

def broadcast_queue_update():
    today = datetime.now().date()
    doctors = Doctor.query.filter_by(is_active=True).all()
    data = []
    for doc in doctors:
        current = Appointment.query.filter_by(doctor_id=doc.id, appointment_date=today, status="called").first()
        waiting_count = Appointment.query.filter_by(doctor_id=doc.id, appointment_date=today, status="waiting").count()
        data.append({"doctor_name": doc.name, "specialization": doc.specialization, "current_queue": current.queue_number if current else "-", "waiting": waiting_count})
    socketio.emit("queue_update", {"doctors": data})

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            flash('Silakan login terlebih dahulu.', 'warning')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def role_required(allowed_roles):
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if 'user_id' not in session or session.get('role') not in allowed_roles:
                flash('Anda tidak memiliki hak akses.', 'danger')
                return redirect(url_for('index'))
            return f(*args, **kwargs)
        return decorated_function
    return decorator

# ── Routes: Auth & Public ────────────────────────────────────
@app.route("/")
def index():
    today = datetime.now().date()
    return render_template("index.html", 
        total_today=Appointment.query.filter_by(appointment_date=today).count(),
        waiting_today=Appointment.query.filter_by(appointment_date=today, status="waiting").count(),
        completed_today=Appointment.query.filter_by(appointment_date=today, status="completed").count(),
        doctors_count=Doctor.query.filter_by(is_active=True).count()
    )

@app.route("/login", methods=["GET", "POST"])
def login():
    if 'user_id' in session: return redirect(url_for('index'))
    if request.method == "POST":
        user = User.query.filter_by(username=request.form.get("username", "").strip()).first()
        if user and check_password_hash(user.password_hash, request.form.get("password", "").strip()):
            session['user_id'], session['username'], session['role'] = user.id, user.username, user.role
            flash(f"Selamat datang, {user.username}!", "success")
            return redirect(url_for("index"))
        flash("Username atau password salah!", "danger")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/booking", methods=["GET", "POST"])
def booking():
    doctors = Doctor.query.filter_by(is_active=True).all()
    if request.method == "POST":
        patient_name = request.form.get("patient_name", "").strip()
        phone = request.form.get("phone", "").strip()
        email = request.form.get("email", "").strip()
        doctor_id = request.form.get("doctor_id", type=int)
        date_str = request.form.get("appointment_date", "")
        
        if not all([patient_name, phone, doctor_id, date_str]):
            flash("Harap isi semua field yang wajib!", "danger"); return redirect(url_for("booking"))
        try: appt_date = datetime.strptime(date_str, "%Y-%m-%d").date()
        except: flash("Format tanggal tidak valid!", "danger"); return redirect(url_for("booking"))
        if appt_date < datetime.now().date():
            flash("Tidak bisa booking untuk tanggal yang sudah lewat!", "danger"); return redirect(url_for("booking"))

        doctor = Doctor.query.get(doctor_id)
        if not doctor: flash("Dokter tidak ditemukan!", "danger"); return redirect(url_for("booking"))

        if Appointment.query.filter_by(doctor_id=doctor_id, appointment_date=appt_date).filter(Appointment.status != "cancelled").count() >= doctor.quota:
            flash(f"Kuota dokter {doctor.name} penuh!", "warning"); return redirect(url_for("booking"))

        patient = Patient.query.filter_by(phone=phone).first()
        if not patient:
            patient = Patient(name=patient_name, phone=phone, email=email, medical_record_number=generate_mr_number())
            db.session.add(patient); db.session.flush()
        else:
            if patient.name != patient_name: patient.name = patient_name
            if email and patient.email != email: patient.email = email
            db.session.commit()

        queue_num = get_next_queue(doctor_id, appt_date)
        appointment = Appointment(patient_id=patient.id, doctor_id=doctor_id, appointment_date=appt_date, queue_number=queue_num, status="waiting")
        db.session.add(appointment); db.session.commit()

        appointment.qr_code = generate_qr_code(appointment.id)
        db.session.commit()

        create_notification(patient.id, f"Booking berhasil! No. Antrian: {queue_num}.")
        flash(f"Booking berhasil! No. Antrian: {queue_num}", "success")
        return redirect(url_for("print_ticket", appt_id=appointment.id))

    return render_template("booking.html", doctors=doctors)

@app.route("/print-ticket/<int:appt_id>")
def print_ticket(appt_id):
    return render_template("print_ticket.html", appt=Appointment.query.get_or_404(appt_id))

# ── Routes: Display & Monitor ────────────────────────────────
@app.route("/display")
def queue_display(): return render_template("queue_display.html")

@app.route("/api/queue-display")
def api_queue_display():
    today = datetime.now().date()
    result = []
    for doc in Doctor.query.filter_by(is_active=True).all():
        current = Appointment.query.filter_by(doctor_id=doc.id, appointment_date=today, status="called").order_by(Appointment.id.desc()).first()
        waiting_list = Appointment.query.filter_by(doctor_id=doc.id, appointment_date=today, status="waiting").order_by(Appointment.queue_number.asc()).all()
        
        if current:
            display_current = current.queue_number
            display_next = waiting_list[0].queue_number if len(waiting_list) > 0 else "-"
        else:
            display_current = waiting_list[0].queue_number if len(waiting_list) > 0 else "-"
            display_next = waiting_list[1].queue_number if len(waiting_list) > 1 else "-"
            
        result.append({"doctor": doc.name, "specialization": doc.specialization, "current": display_current, "next": display_next})
    return jsonify(result)

@app.route("/monitor")
def monitor(): return render_template("monitor.html")

@app.route("/api/monitor/<mrn>")
def api_monitor(mrn):
    patient = Patient.query.filter_by(medical_record_number=mrn).first()
    if not patient: return jsonify({"error": "Tidak ditemukan"}), 404
    today = datetime.now().date()
    appt = Appointment.query.filter_by(patient_id=patient.id, appointment_date=today).first()
    if not appt: return jsonify({"status": "Tidak ada antrian", "queue_number": "-", "estimated_wait": "-"})
    ahead = Appointment.query.filter_by(doctor_id=appt.doctor_id, appointment_date=today, status="waiting").filter(Appointment.queue_number < appt.queue_number).count()
    return jsonify({"status": appt.status, "queue_number": appt.queue_number, "position": ahead + 1, "doctor": appt.doctor.name, "estimated_wait": f"{ahead * 10} menit"})

# ── Routes: Admin Area ───────────────────────────────────────
@app.route("/doctors", methods=["GET", "POST"])
@login_required
@role_required(['admin'])
def doctor_schedule():
    if request.method == "POST":
        action = request.form.get("action")
        if action == "add":
            name, spec = request.form.get("name", "").strip(), request.form.get("specialization", "").strip()
            fee = request.form.get("consultation_fee", 150000, type=float)
            schedule = {d: f"{request.form.get('hours_start')}-{request.form.get('hours_end')}" for d in request.form.getlist("days")}
            db.session.add(Doctor(name=name, specialization=spec, schedule=json.dumps(schedule), quota=request.form.get("quota", 30, type=int), consultation_fee=fee))
            db.session.commit(); flash("Dokter ditambahkan!", "success")
        elif action == "delete":
            doc = Doctor.query.get(request.form.get("doctor_id", type=int))
            if doc: db.session.delete(doc); db.session.commit()
        return redirect(url_for("doctor_schedule"))
    return render_template("doctor_schedule.html", doctors=Doctor.query.all())

@app.route("/admin/queue", methods=["GET", "POST"])
@login_required
def admin_queue():
    today = datetime.now().date()
    if request.method == "POST":
        appt = Appointment.query.get(request.form.get("appointment_id", type=int))
        action = request.form.get("action")
        if appt:
            if action == "call":
                Appointment.query.filter_by(doctor_id=appt.doctor_id, appointment_date=today, status="called").update({"status": "waiting"})
                appt.status = "called"; db.session.commit()
                create_notification(appt.patient_id, f"No. {appt.queue_number} DIPANGGIL!")
                broadcast_queue_update()
                socketio.emit("announce_call", {"queue_number": appt.queue_number, "doctor_name": appt.doctor.name})
            
            elif action == "start":
                appt.status = "in_progress"
                db.session.commit()
                broadcast_queue_update()
                
            elif action == "complete":
                appt.status = "completed"; db.session.commit()
                invoice = Invoice(appointment_id=appt.id, patient_id=appt.patient_id, consultation_fee=appt.doctor.consultation_fee, total=appt.doctor.consultation_fee, status="unpaid")
                db.session.add(invoice); db.session.commit()
                broadcast_queue_update()
            elif action == "cancel":
                appt.status = "cancelled"; db.session.commit()
        return redirect(url_for("admin_queue"))
    return render_template("queue_display.html", appointments=Appointment.query.filter_by(appointment_date=today).all(), admin=True)

@app.route("/billing", methods=["GET", "POST"])
@login_required
def billing():
    if request.method == "POST":
        invoice_id = request.form.get("invoice_id", type=int)
        action = request.form.get("action")
        invoice = Invoice.query.get(invoice_id)
        if invoice:
            if action == "pay":
                invoice.status = "paid"
                db.session.commit()
                flash("Pembayaran berhasil!", "success")
            elif action == "add_item":
                desc = request.form.get("description", "").strip()
                qty = request.form.get("quantity", 1, type=int)
                price = request.form.get("price", 0, type=float)
                db.session.add(InvoiceItem(
                    invoice_id=invoice.id, 
                    description=desc, 
                    quantity=qty, 
                    price=price, 
                    total=qty*price
                ))
                invoice.medicine_fee += (qty * price)
                invoice.total = invoice.consultation_fee + invoice.medicine_fee
                db.session.commit()
                flash("Obat ditambahkan!", "success")
        return redirect(url_for("billing"))
    
    medicines = Medicine.query.filter(Medicine.stock > 0).all()
    invoices = Invoice.query.order_by(Invoice.created_at.desc()).limit(50).all()
    return render_template("billing.html", invoices=invoices, medicines=medicines)

@app.route("/print-invoice/<int:invoice_id>")
@login_required
def print_invoice(invoice_id):
    invoice = Invoice.query.get_or_404(invoice_id)
    patient = Patient.query.get(invoice.patient_id)
    appointment = Appointment.query.get(invoice.appointment_id)
    doctor = Doctor.query.get(appointment.doctor_id) if appointment else None
    return render_template("print_invoice.html", invoice=invoice, patient=patient, appointment=appointment, doctor=doctor)

@app.route("/pharmacy", methods=["GET", "POST"])
@login_required
@role_required(['admin'])
def pharmacy():
    if request.method == "POST":
        action = request.form.get("action")
        if action == "add":
            db.session.add(Medicine(name=request.form.get("name", "").strip(), category=request.form.get("category", "Umum"), stock=request.form.get("stock", 0, type=int), price=request.form.get("price", 0, type=float), unit=request.form.get("unit", "tablet")))
            db.session.commit(); flash("Obat ditambahkan!", "success")
        elif action == "update_stock":
            med = Medicine.query.get(request.form.get("medicine_id", type=int))
            if med:
                med.stock += request.form.get("quantity", 0, type=int)
                db.session.commit(); flash("Stok diupdate!", "success")
        return redirect(url_for("pharmacy"))
    return render_template("pharmacy.html", medicines=Medicine.query.all(), low_stock=Medicine.query.filter(Medicine.stock <= Medicine.min_stock).all())

# ═══════════════════════════════════════════════════════════════
# ✅ PERBAIKAN: Route pharmacy-queue sekarang mengirim data 'doctors'
# ═══════════════════════════════════════════════════════════════
@app.route("/pharmacy-queue")
@login_required
def pharmacy_queue():
    """Halaman display antrian farmasi untuk mengambil obat dengan filter dokter"""
    today = datetime.now().date()
    
    # Ambil semua dokter aktif untuk ditampilkan sebagai tombol filter
    doctors = Doctor.query.filter_by(is_active=True).all()
    
    # Ambil semua appointment yang sudah selesai dan lunas
    completed_appts = Appointment.query.filter_by(
        appointment_date=today, 
        status="completed"
    ).join(Invoice).filter(Invoice.status == "paid").order_by(Appointment.queue_number.asc()).all()
    
    return render_template("pharmacy_queue.html", 
                          appointments=completed_appts, 
                          doctors=doctors)
# ═══════════════════════════════════════════════════════════════

@app.route("/api/pharmacy-queue")
def api_pharmacy_queue():
    today = datetime.now().date()
    completed_appts = Appointment.query.filter_by(appointment_date=today, status="completed").join(Invoice).filter(Invoice.status == "paid").order_by(Appointment.queue_number.asc()).all()
    result = []
    for appt in completed_appts:
        result.append({"id": appt.id, "queue_number": appt.queue_number, "patient_name": appt.patient.name, "doctor_name": appt.doctor.name, "pharmacy_status": appt.pharmacy_status})
    return jsonify(result)

@app.route("/pharmacy/call/<int:appt_id>", methods=["POST"])
@login_required
def pharmacy_call(appt_id):
    appt = Appointment.query.get_or_404(appt_id)
    appt.pharmacy_status = "called"
    db.session.commit()
    socketio.emit("pharmacy_call", {"queue_number": appt.queue_number, "patient_name": appt.patient.name, "doctor_name": appt.doctor.name})
    flash(f"No. Antrian {appt.queue_number} ({appt.patient.name}) dipanggil untuk ambil obat!", "success")
    return redirect(url_for("pharmacy_queue"))

@app.route("/medical-records", methods=["GET", "POST"])
@login_required
def medical_records():
    if request.method == "POST":
        patient = Patient.query.filter_by(medical_record_number=request.form.get("patient_mrn", "").strip()).first()
        if not patient: flash("Pasien tidak ditemukan!", "danger"); return redirect(url_for("medical_records"))
        db.session.add(MedicalRecord(patient_id=patient.id, appointment_id=request.form.get("appointment_id", type=int) or None, diagnosis=request.form.get("diagnosis", "").strip(), treatment=request.form.get("treatment", "").strip(), notes=request.form.get("notes", "").strip()))
        db.session.commit()
        create_notification(patient.id, "Rekam medis diperbarui.")
        flash("Rekam medis disimpan!", "success"); return redirect(url_for("medical_records"))
    return render_template("medical_records.html", records=MedicalRecord.query.order_by(MedicalRecord.created_at.desc()).limit(50).all(), patients=Patient.query.all())

@app.route("/admin/feedbacks")
@login_required
@role_required(['admin'])
def admin_feedbacks():
    feedbacks = Feedback.query.order_by(Feedback.created_at.desc()).all()
    avg_rating = db.session.query(db.func.avg(Feedback.rating)).scalar() or 0
    return render_template("admin_feedbacks.html", feedbacks=feedbacks, avg_rating=round(avg_rating, 1))

@app.route("/admin/reports")
@login_required
@role_required(['admin'])
def reports():
    today = datetime.now().date()
    total = Appointment.query.filter_by(appointment_date=today).count()
    completed = Appointment.query.filter_by(appointment_date=today, status="completed").count()
    revenue = db.session.query(db.func.sum(Invoice.total)).filter(Invoice.status == "paid").scalar() or 0
    doctor_stats = db.session.query(Doctor.name, db.func.count(Appointment.id)).join(Appointment).filter(Appointment.appointment_date == today).group_by(Doctor.name).all()
    return render_template("reports.html", total=total, completed=completed, revenue=revenue, doctor_stats=doctor_stats)

@app.route("/export/excel")
@login_required
@role_required(['admin'])
def export_excel():
    wb = Workbook()
    ws = wb.active
    ws.append(["Tanggal", "No. Antrian", "Pasien", "Dokter", "Status", "Total"])
    for appt in Appointment.query.order_by(Appointment.appointment_date.desc()).all():
        inv = appt.invoice
        ws.append([appt.appointment_date.strftime("%d/%m/%Y"), appt.queue_number, appt.patient.name, appt.doctor.name, appt.status, inv.total if inv else 0])
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name=f"laporan.xlsx")

@app.route("/notifications/<mrn>")
def notifications(mrn):
    patient = Patient.query.filter_by(medical_record_number=mrn).first()
    if not patient:
        flash("Pasien tidak ditemukan!", "danger")
        return redirect(url_for("index"))
    notifs = Notification.query.filter_by(patient_id=patient.id).order_by(Notification.created_at.desc()).all()
    return render_template("notifications.html", patient=patient, notifications=notifs)

@app.route("/api/notifications/<mrn>")
def api_notifications(mrn):
    patient = Patient.query.filter_by(medical_record_number=mrn).first()
    if not patient: return jsonify({"error": "Not found"}), 404
    notifs = Notification.query.filter_by(patient_id=patient.id).order_by(Notification.created_at.desc()).limit(20).all()
    return jsonify([{"message": n.message, "is_read": n.is_read, "time": n.created_at.strftime("%H:%M:%S")} for n in notifs])

@app.route("/feedback-access", methods=["GET", "POST"])
def feedback_access():
    patient = None
    completed_appts = []
    if request.method == "POST":
        mrn = request.form.get("mrn", "").strip().upper()
        patient = Patient.query.filter_by(medical_record_number=mrn).first()
        if patient:
            completed_appts = Appointment.query.filter_by(patient_id=patient.id, status="completed").order_by(Appointment.appointment_date.desc()).all()
        else:
            flash("No. Rekam Medis tidak ditemukan!", "danger")
    return render_template("feedback_access.html", patient=patient, appointments=completed_appts)

@app.route("/feedback/<int:appt_id>", methods=["GET", "POST"])
def feedback_form(appt_id):
    appointment = Appointment.query.get_or_404(appt_id)
    if appointment.status != 'completed':
        flash("Feedback hanya bisa diberikan setelah kunjungan selesai.", "warning")
        return redirect(url_for("feedback_access"))
    already_feedback = Feedback.query.filter_by(appointment_id=appt_id).first()
    if request.method == "POST":
        rating = request.form.get("rating", type=int)
        comment = request.form.get("comment", "").strip()
        if not rating or rating < 1 or rating > 5:
            flash("Rating harus antara 1-5!", "danger")
            return redirect(url_for("feedback_form", appt_id=appt_id))
        if already_feedback:
            already_feedback.rating, already_feedback.comment = rating, comment
            flash("Feedback berhasil diperbarui! Terima kasih.", "success")
        else:
            db.session.add(Feedback(appointment_id=appt_id, patient_id=appointment.patient_id, rating=rating, comment=comment))
            flash("Terima kasih atas feedback Anda! 🌟", "success")
        db.session.commit()
        return redirect(url_for("feedback_access"))
    return render_template("feedback.html", appointment=appointment, already_feedback=already_feedback)

# ═══════════════════════════════════════════════════════════════
# WebSocket & Scheduler
# ══════════════════════════════════════════════════════════════
@socketio.on("connect")
def handle_connect(): 
    emit("server_message", {"msg": "Terhubung"})

def send_reminders():
    tomorrow = datetime.now().date() + timedelta(days=1)
    upcoming = Appointment.query.filter_by(appointment_date=tomorrow, status="waiting").all()
    for appt in upcoming:
        create_notification(appt.patient_id, f"Pengingat: Jadwal dengan Dr. {appt.doctor.name} besok, No. Antrian: {appt.queue_number}.")
    if upcoming: print(f"[Scheduler] {len(upcoming)} pengingat dikirim.")

def seed_data():
    if User.query.count() == 0:
        db.session.add(User(username="admin", password_hash=generate_password_hash("admin123"), role="admin"))
        db.session.commit()
    if Doctor.query.count() == 0:
        db.session.add(Doctor(name="Dr. Andi", specialization="Umum", schedule=json.dumps({"Senin": "08:00-14:00"}), quota=30, consultation_fee=150000))
        db.session.commit()
    if Medicine.query.count() == 0:
        db.session.add(Medicine(name="Paracetamol 500mg", category="Obat Bebas", stock=100, price=5000, unit="tablet"))
        db.session.commit()

if __name__ == "__main__":
    with app.app_context():
        db.create_all()
        seed_data()
    
    scheduler = BackgroundScheduler()
    scheduler.add_job(send_reminders, "cron", hour=18, minute=0)
    scheduler.start()
    
    print("=" * 60)
    print("  🏥 AULIA KLINIK - V3.0 ENTERPRISE EDITION")
    print("  🌐 http://127.0.0.1:5001 | admin/admin123")
    print("=" * 60)
    socketio.run(app, host="0.0.0.0", port=5001, debug=True, allow_unsafe_werkzeug=True)