# Import date and datetime to handle dates and timestamps
from datetime import date, datetime
# Import pandas to read/write the Excel file with dummy data and appended records
import pandas as pd

# Define a constant for the Excel file name which stores dummy and new data
EXCEL_FILE = "hmis_dummy_data.xlsx"

# Create a list of doctors with IDs, names, and departments
DOCTORS = [
    {"id": "DR001", "name": "Dr. Arjun Menon",   "department": "General Medicine"},
    {"id": "DR002", "name": "Dr. Meera Rao",     "department": "General Surgery"},
    {"id": "DR003", "name": "Dr. Rohan Kulkarni","department": "Paediatrics"},
    {"id": "DR004", "name": "Dr. Kavita Nair",   "department": "Ophthalmology"},
    {"id": "DR005", "name": "Dr. Sanjay Patel",  "department": "ENT"},
    {"id": "DR006", "name": "Dr. Ananya Sharma", "department": "Obstetrics & Gynaecology"},
    {"id": "DR007", "name": "Dr. Vivek Reddy",   "department": "Orthopaedics"},
    {"id": "DR008", "name": "Dr. Neha Gupta",    "department": "Dermatology"},
    {"id": "DR009", "name": "Dr. Sameer Iyer",   "department": "Psychiatry"},
    {"id": "DR010", "name": "Dr. Priya Das",     "department": "General OPD"},
]

# Ordered list of departments
DEPARTMENTS = [
    "General Medicine",
    "General Surgery",
    "Paediatrics",
    "Ophthalmology",
    "ENT",
    "Obstetrics & Gynaecology",
    "Orthopaedics",
    "Dermatology",
    "Psychiatry",
    "General OPD",
]


# -------------------------
# Helper functions
# -------------------------

# Function to keep only digit characters from any input
def clean_digits(value):
    s = str(value)
    return "".join(ch for ch in s if ch.isdigit())

# Function to generate UHID: first 4 digits of DOB (DDMM) + last 4 digits of Aadhaar
def generate_uhid(dob_digits_8, aadhaar_digits_12):
    first_four = dob_digits_8[:4]           # DDMM from DOB
    last_four = aadhaar_digits_12[-4:]      # last 4 from Aadhaar
    return first_four + last_four           # 8-digit UHID

# Function to parse DDMMYYYY into a date object
def parse_dob_from_digits(dob_digits_8):
    day = int(dob_digits_8[0:2])
    month = int(dob_digits_8[2:4])
    year = int(dob_digits_8[4:8])
    return date(year, month, day)

# Validate Aadhaar from user (12 digits, numeric)
def get_valid_aadhaar_digits():
    while True:
        aadhaar_input = input("Enter Aadhaar number (12 digits, numbers only): ").strip()
        if not aadhaar_input.isdigit():
            print("Error: Aadhaar must contain only numbers. Please try again.")
            continue
        if len(aadhaar_input) != 12:
            print("Error: Aadhaar must be exactly 12 digits. Please try again.")
            continue
        return aadhaar_input

# Validate DOB from user and return (date, "ddmmyyyy" string)
def get_valid_dob_date_and_digits():
    while True:
        dob_input = input("Enter DOB as DDMMYYYY (8 digits, numbers only): ").strip()
        if not dob_input.isdigit():
            print("Error: DOB must contain only numbers. Please try again.")
            continue
        if len(dob_input) != 8:
            print("Error: DOB must be exactly 8 digits (DDMMYYYY). Please try again.")
            continue
        try:
            dob_date = parse_dob_from_digits(dob_input)
            return dob_date, dob_input
        except Exception:
            print("Error: Invalid date. Please enter a valid DOB in DDMMYYYY format.")
            continue

# Validate UHID from user (8 digits, numeric)
def get_valid_uhid_input():
    while True:
        uhid_input = input("Enter UHID (8 digits, numbers only): ").strip()
        if not uhid_input.isdigit():
            print("Error: UHID must contain only numbers. Please try again.")
            continue
        if len(uhid_input) != 8:
            print("Error: UHID must be exactly 8 digits. Please try again.")
            continue
        return uhid_input

# Fetch the doctor record for a particular department
def get_doctor_for_department(department):
    for doctor in DOCTORS:
        if doctor["department"] == department:
            return doctor
    return None

# Let user select department and auto-assign relevant doctor
def select_department_and_doctor():
    print("\nSelect Department:")
    for idx, dept in enumerate(DEPARTMENTS, start=1):
        print(f"{idx}. {dept}")
    while True:
        choice = input("Enter department number: ").strip()
        if not choice.isdigit():
            print("Error: Please enter a number for the department.")
            continue
        choice_num = int(choice)
        if choice_num < 1 or choice_num > len(DEPARTMENTS):
            print("Error: Invalid department choice. Please try again.")
            continue
        department = DEPARTMENTS[choice_num - 1]
        doctor = get_doctor_for_department(department)
        print(f"Assigned Doctor: {doctor['id']} - {doctor['name']} ({doctor['department']})")
        return department, doctor["id"]


# -------------------------
# Core classes
# -------------------------

# Class to store investigation data (no result)
class Investigation:
    def __init__(self, inv_type, name):
        self.inv_type = inv_type
        self.name = name
        self.ordered_on = datetime.now()

# Class to store medication data
class Medication:
    def __init__(self, name, dose, frequency, duration, instructions=""):
        self.name = name
        self.dose = dose
        self.frequency = frequency
        self.duration = duration
        self.instructions = instructions

# Class to store vitals
class Vitals:
    def __init__(self, bp=None, pulse=None, spo2=None, temperature=None):
        self.bp = bp
        self.pulse = pulse
        self.spo2 = spo2
        self.temperature = temperature

# Class to store clinical history
class ClinicalHistory:
    def __init__(
        self,
        presenting_complaints="",
        past_history="",
        surgical_history="",
        allergy_history="",
        medication_history="",
        addiction_history=""
    ):
        self.presenting_complaints = presenting_complaints
        self.past_history = past_history
        self.surgical_history = surgical_history
        self.allergy_history = allergy_history
        self.medication_history = medication_history
        self.addiction_history = addiction_history

# Class to store visit details
class Visit:
    def __init__(self, visit_id, patient_uhid, department, doctor_id):
        self.visit_id = visit_id
        self.patient_uhid = patient_uhid
        self.department = department
        self.doctor_id = doctor_id
        self.created_on = datetime.now()
        self.vitals = Vitals()
        self.clinical_history = ClinicalHistory()
        self.diagnosis = ""
        self.investigations = []
        self.medications = []
        self.advice = ""
        self.disposition = ""
        self.is_closed = False

    def add_vitals(self, vitals):
        self.vitals = vitals

    def add_clinical_history(self, history):
        self.clinical_history = history

    def set_diagnosis(self, diagnosis):
        self.diagnosis = diagnosis

    def add_investigation(self, investigation):
        self.investigations.append(investigation)

    def add_medication(self, medication):
        self.medications.append(medication)

    def set_advice(self, advice, disposition=""):
        self.advice = advice
        if disposition:
            self.disposition = disposition

    def close_visit(self):
        self.is_closed = True

    def generate_case_sheet(self):
        lines = []
        lines.append(f"VISIT ID: {self.visit_id}")
        lines.append(f"UHID: {self.patient_uhid}")
        lines.append(f"Department: {self.department}")
        lines.append(f"Doctor: {self.doctor_id}")
        lines.append(f"Visit Date: {self.created_on.strftime('%Y-%m-%d %H:%M')}")
        lines.append("\n--- VITALS ---")
        lines.append(
            f"BP: {self.vitals.bp}, Pulse: {self.vitals.pulse}, "
            f"SpO2: {self.vitals.spo2}, Temp: {self.vitals.temperature}"
        )
        lines.append("\n--- CLINICAL HISTORY ---")
        lines.append(f"Presenting Complaints: {self.clinical_history.presenting_complaints}")
        lines.append(f"Past History: {self.clinical_history.past_history}")
        lines.append(f"Surgical History: {self.clinical_history.surgical_history}")
        lines.append(f"Allergy History: {self.clinical_history.allergy_history}")
        lines.append(f"Medication History: {self.clinical_history.medication_history}")
        lines.append(f"Addiction History: {self.clinical_history.addiction_history}")
        lines.append("\n--- DIAGNOSIS ---")
        lines.append(self.diagnosis or "")
        lines.append("\n--- INVESTIGATIONS ---")
        if not self.investigations:
            lines.append("None")
        else:
            for inv in self.investigations:
                lines.append(
                    f"{inv.inv_type} - {inv.name} "
                    f"(Ordered: {inv.ordered_on})"
                )
        lines.append("\n--- MEDICATIONS ---")
        if not self.medications:
            lines.append("None")
        else:
            for med in self.medications:
                lines.append(
                    f"{med.name} {med.dose}, {med.frequency} for {med.duration}. {med.instructions}"
                )
        lines.append("\n--- ADVICE & DISPOSITION ---")
        lines.append(self.advice or "")
        lines.append(f"Disposition: {self.disposition or 'Not recorded'}")
        lines.append(f"Visit Closed: {self.is_closed}")
        return "\n".join(lines)


# Class to store patient details (with private attributes)
class Patient:
    def __init__(self, uhid, name, gender, aadhaar_digits, dob, occupation):
        self.__uhid = uhid                                # private attribute UHID
        self.__name = name                                # private attribute Name
        self.gender = gender
        self.__aadhaar_digits = aadhaar_digits            # private attribute Aadhaar number
        self.__dob = dob                                  # private attribute DOB
        self.occupation = occupation
        self.address_line1 = ""
        self.city = ""
        self.state = ""
        self.pincode = ""
        self.address = ""
        self.mobile = ""
        self.emergency_name = ""
        self.emergency_relation = ""
        self.emergency_phone = ""
        self.visits = []

    def get_uhid(self):                                    # getter for UHID
        return self.__uhid

    def get_name(self):                                    # getter for Name
        return self.__name

    def get_aadhaar_digits(self):                          # getter for Aadhaar number
        return self.__aadhaar_digits

    def get_dob(self):                                     # getter for DOB
        return self.__dob

    @property
    def age(self):
        if not self.__dob:
            return None
        today = date.today()
        years = today.year - self.__dob.year
        if (today.month, today.day) < (self.__dob.month, self.__dob.day):
            years -= 1
        return years

    def add_visit(self, visit):
        self.visits.append(visit)


# Main system class for HMIS OPD
class HMISOPDSystem:
    def __init__(self):
        self.patients = {}  # UHID -> Patient object

    def register_patient(
        self,
        name,
        gender,
        aadhaar_digits,
        dob,
        occupation,
        address_line1,
        city,
        state,
        pincode,
        mobile,
        emergency_name,
        emergency_relation,
        emergency_phone,
        uhid=None,       # optional UHID (for Excel data)
    ):
        # If UHID is not provided (front-desk registration), generate from DOB + Aadhaar
        if uhid is None:
            dob_digits_8 = dob.strftime("%d%m%Y")
            uhid = generate_uhid(dob_digits_8, aadhaar_digits)

        # Clean and validate UHID as 8-digit string
        uhid = clean_digits(uhid)
        if len(uhid) != 8:
            raise ValueError(f"Invalid UHID computed or provided: {uhid}")

        # If patient already exists, return existing
        if uhid in self.patients:
            return self.patients[uhid]

        patient = Patient(uhid, name, gender, aadhaar_digits, dob, occupation)
        patient.address_line1 = address_line1
        patient.city = city
        patient.state = state
        patient.pincode = pincode
        patient.address = f"{address_line1}, {city}, {state} - {pincode}"
        patient.mobile = mobile
        patient.emergency_name = emergency_name
        patient.emergency_relation = emergency_relation
        patient.emergency_phone = emergency_phone

        self.patients[uhid] = patient
        return patient

    def create_visit(self, patient_uhid, department, doctor_id):
        if patient_uhid not in self.patients:
            raise ValueError("Patient not found")
        visit_id = f"VISIT-{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
        visit = Visit(visit_id, patient_uhid, department, doctor_id)
        self.patients[patient_uhid].add_visit(visit)
        return visit

    def get_open_visits(self, patient_uhid):
        if patient_uhid not in self.patients:
            return []
        return [v for v in self.patients[patient_uhid].visits if not v.is_closed]

    def get_visit(self, patient_uhid, visit_id):
        if patient_uhid not in self.patients:
            raise ValueError("Patient not found")
        for v in self.patients[patient_uhid].visits:
            if v.visit_id == visit_id:
                return v
        raise ValueError("Visit not found")

    # Load dummy data from Excel using UHID from sheet
    def load_dummy_data_from_excel(self, excel_path):
        try:
            # Read as strings so ddmmyyyy and IDs preserve leading zeros
            df = pd.read_excel(excel_path, dtype=str)
        except FileNotFoundError:
            print(f"Warning: Excel file '{excel_path}' not found. Starting with empty system.")
            return

        loaded_count = 0

        for _, row in df.iterrows():
            # UHID from Excel (must be 8 digits)
            raw_uhid = (row.get("UHID") or "").strip()
            uhid_digits = clean_digits(raw_uhid)
            if len(uhid_digits) != 8:
                continue

            # DOB from Excel (assumed ddmmyyyy string)
            raw_dob = (row.get("DOB") or "").strip()
            dob_digits = clean_digits(raw_dob)
            if len(dob_digits) != 8:
                continue
            try:
                dob = parse_dob_from_digits(dob_digits)
            except Exception:
                continue

            # Aadhaar from Excel (12 digits)
            raw_aadhaar = (row.get("Aadhaar") or "").strip()
            aadhaar_digits = clean_digits(raw_aadhaar)
            if len(aadhaar_digits) != 12:
                continue

            # Other patient details
            name = str(row.get("Name", "")).strip()
            gender = str(row.get("Gender", "")).strip()
            occupation = str(row.get("Occupation", "")).strip()
            addr1 = str(row.get("AddressLine1", "")).strip()
            city = str(row.get("City", "")).strip()
            state = str(row.get("State", "")).strip()
            pincode = str(row.get("Pincode", "")).strip()
            mobile = str(row.get("Mobile", "")).strip()
            emergency_name = str(row.get("EmergencyName", "")).strip()
            emergency_relation = str(row.get("EmergencyRelation", "")).strip()
            emergency_phone = str(row.get("EmergencyPhone", "")).strip()

            # Register patient using UHID from Excel
            patient = self.register_patient(
                name=name,
                gender=gender,
                aadhaar_digits=aadhaar_digits,
                dob=dob,
                occupation=occupation,
                address_line1=addr1,
                city=city,
                state=state,
                pincode=pincode,
                mobile=mobile,
                emergency_name=emergency_name,
                emergency_relation=emergency_relation,
                emergency_phone=emergency_phone,
                uhid=uhid_digits,
            )
            uhid = patient.get_uhid()

            # Department & Doctor
            department = str(row.get("Department", "General Medicine")).strip()
            doctor_id = str(row.get("DoctorId", "")).strip()
            if doctor_id not in [d["id"] for d in DOCTORS]:
                doctor = get_doctor_for_department(department)
                doctor_id = doctor["id"] if doctor else "DR001"

            visit = self.create_visit(uhid, department, doctor_id)

            # Presenting complaints / history
            presenting = str(row.get("PresentingComplaints", "")).strip()
            if presenting:
                history = ClinicalHistory(presenting_complaints=presenting)
                visit.add_clinical_history(history)

            # Investigations as free text
            inv_text = str(row.get("Investigations", "")).strip()
            if inv_text:
                visit.add_investigation(Investigation("Text", inv_text))

            # Medications as free text
            med_text = str(row.get("Medications", "")).strip()
            if med_text:
                visit.add_medication(Medication(med_text, "", "", ""))

            diagnosis = str(row.get("Diagnosis", "")).strip()
            disposition = str(row.get("Disposition", "Follow up")).strip()
            advice_text = str(row.get("Advice", "")).strip()
            if not advice_text and diagnosis:
                advice_text = "Review after 5 days"

            visit.set_diagnosis(diagnosis)
            visit.set_advice(advice_text, disposition=disposition)
            visit.close_visit()

            loaded_count += 1


    # Append a completed visit to Excel including Medications, Advice, etc.
    def append_visit_to_excel(self, excel_path, patient, visit):
        df = pd.read_excel(excel_path, dtype=str)

        # Flatten investigations list (no result)
        if not visit.investigations:
            inv_text = ""

        else:
            inv_parts = []
            for inv in visit.investigations:
                inv_parts.append(f"{inv.inv_type} - {inv.name}")
            inv_text = "; ".join(inv_parts)

        # Flatten medications list
        if not visit.medications:
            med_text = ""
        else:
            med_parts = []
            for med in visit.medications:
                med_parts.append(f"{med.name} {med.dose}, {med.frequency} for {med.duration}")
            med_text = "; ".join(med_parts)

        dob_obj = patient.get_dob()
        dob_str = dob_obj.strftime("%d%m%Y") if isinstance(dob_obj, date) else ""

        row = {
            "UHID": patient.get_uhid(),
            "Name": patient.get_name(),
            "Gender": patient.gender,
            "DOB": dob_str,
            "Aadhaar": patient.get_aadhaar_digits(),
            "Occupation": patient.occupation,
            "AddressLine1": patient.address_line1,
            "City": patient.city,
            "State": patient.state,
            "Pincode": patient.pincode,
            "Mobile": patient.mobile,
            "EmergencyName": patient.emergency_name,
            "EmergencyRelation": patient.emergency_relation,
            "EmergencyPhone": patient.emergency_phone,
            "Department": visit.department,
            "DoctorId": visit.doctor_id,
            "PresentingComplaints": visit.clinical_history.presenting_complaints,
            "Investigations": inv_text,
            "Medications": med_text,
            "Advice": visit.advice,
            "Diagnosis": visit.diagnosis,
            "Disposition": visit.disposition,
        }

        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        df.to_excel(excel_path, index=False)


# -------------------------
# Console flows
# -------------------------

# Registration flow at front desk
def registration_flow(system):
    print("\n--- Patient Registration Desk ---")
    # Ask if existing UHID with strict y/n
    while True:
        has_uhid = input("Does the patient already have a UHID? (y/n): ").strip().lower()
        if has_uhid in ("y", "n"):
            break
        print("Error: Please enter 'y' or 'n'.")

    if has_uhid == "y":
        uhid = get_valid_uhid_input()
        if uhid not in system.patients:
            print("Error: UHID not found in the system.")
            return
        department, doctor_id = select_department_and_doctor()
        visit = system.create_visit(uhid, department, doctor_id)
        print(f"Linked existing UHID {uhid} to new visit {visit.visit_id} (Pending doctor consultation).")
        return

    # New patient branch
    name = input("Patient Name: ").strip()
    gender = input("Gender (Male/Female/Other): ").strip()
    dob, dob_digits_8 = get_valid_dob_date_and_digits()
    aadhaar_digits = get_valid_aadhaar_digits()
    occupation = input("Occupation: ").strip()
    addr1 = input("Address Line 1: ").strip()
    city = input("City: ").strip()
    state = input("State: ").strip()
    pincode = input("Pincode: ").strip()
    mobile = input("Mobile: ").strip()
    emergency_name = input("Emergency Contact Name: ").strip()
    emergency_relation = input("Emergency Contact Relationship: ").strip()
    emergency_phone = input("Emergency Contact Phone: ").strip()

    patient = system.register_patient(
        name=name,
        gender=gender,
        aadhaar_digits=aadhaar_digits,
        dob=dob,
        occupation=occupation,
        address_line1=addr1,
        city=city,
        state=state,
        pincode=pincode,
        mobile=mobile,
        emergency_name=emergency_name,
        emergency_relation=emergency_relation,
        emergency_phone=emergency_phone,
    )

    department, doctor_id = select_department_and_doctor()
    uhid = patient.get_uhid()
    visit = system.create_visit(uhid, department, doctor_id)
    print(f"New patient registered. UHID: {patient.get_uhid()}")
    print(f"Visit created and sent for doctor consultation. VISIT ID: {visit.visit_id}")


# Doctor consultation flow
def doctor_consultation_flow(system):
    print("\n--- Doctor Consultation ---")
    uhid = get_valid_uhid_input()
    if uhid not in system.patients:
        print("Error: Patient not found.")
        return

    open_visits = system.get_open_visits(uhid)
    if not open_visits:
        print("No open/pending visits for this UHID.")
        return

    if len(open_visits) > 1:
        print("Open visits:")
        for idx, v in enumerate(open_visits, start=1):
            print(f"{idx}. {v.visit_id} | Dept: {v.department} | Doctor: {v.doctor_id} | Created: {v.created_on}")
        choice = input("Select visit number: ").strip()
        if not choice.isdigit():
            print("Error: Invalid choice.")
            return
        choice_num = int(choice)
        if choice_num < 1 or choice_num > len(open_visits):
            print("Error: Invalid choice.")
            return
        visit = open_visits[choice_num - 1]
    else:
        visit = open_visits[0]

    print(f"Using VISIT ID: {visit.visit_id}")

    bp = input("BP (e.g. 120/80): ").strip() or None
    pulse_input = input("Pulse: ").strip()
    spo2_input = input("SpO2: ").strip()
    temp_input = input("Temperature: ").strip()
    pulse = int(pulse_input) if pulse_input.isdigit() else None
    spo2 = int(spo2_input) if spo2_input.isdigit() else None
    try:
        temperature = float(temp_input) if temp_input else None
    except ValueError:
        temperature = None
    vitals = Vitals(bp=bp, pulse=pulse, spo2=spo2, temperature=temperature)
    visit.add_vitals(vitals)

    presenting_complaints = input("Presenting Complaints: ").strip()
    past_history = input("Past History: ").strip()
    surgical_history = input("Surgical History: ").strip()
    allergy_history = input("Allergy History: ").strip()
    medication_history = input("Medication History: ").strip()
    addiction_history = input("Addiction History: ").strip()
    history = ClinicalHistory(
        presenting_complaints=presenting_complaints,
        past_history=past_history,
        surgical_history=surgical_history,
        allergy_history=allergy_history,
        medication_history=medication_history,
        addiction_history=addiction_history,
    )
    visit.add_clinical_history(history)

    diagnosis = input("Diagnosis: ").strip()
    visit.set_diagnosis(diagnosis)

    # Add investigations (no result capture)
    while True:
        add_inv = input("Add investigation? (y/n): ").strip().lower()
        if add_inv in ("y", "n"):
            if add_inv == "n":
                break
            inv_type = input("  Investigation Type (e.g. Blood, X-Ray): ").strip()
            inv_name = input("  Investigation Name (e.g. CBC): ").strip()
            investigation = Investigation(inv_type, inv_name)
            visit.add_investigation(investigation)
        else:
            print("Error: Please enter 'y' or 'n'.")

    # Add medications
    while True:
        add_med = input("Add medication? (y/n): ").strip().lower()
        if add_med in ("y", "n"):
            if add_med == "n":
                break
            med_name = input("  Medicine Name: ").strip()
            med_dose = input("  Dose: ").strip()
            med_frequency = input("  Frequency: ").strip()
            med_duration = input("  Duration: ").strip()
            med_instructions = input("  Instructions: ").strip()
            medication = Medication(med_name, med_dose, med_frequency, med_duration, med_instructions)
            visit.add_medication(medication)
        else:
            print("Error: Please enter 'y' or 'n'.")

    advice = input("Advice / Instructions to patient: ").strip()
    disposition = input("Disposition (Follow up / Referred / Discharged): ").strip()
    visit.set_advice(advice, disposition=disposition)
    visit.close_visit()
    print("\nConsultation saved and visit closed.")
    print("\n--- OPD Case Sheet Preview ---")
    print(visit.generate_case_sheet())

    patient = system.patients[uhid]
    system.append_visit_to_excel(EXCEL_FILE, patient, visit)


# View patient history flow
def view_history_flow(system):
    print("\n--- View Patient History ---")
    uhid = get_valid_uhid_input()
    if uhid not in system.patients:
        print("Error: Patient not found.")
        return

    patient = system.patients[uhid]
    print(f"Patient: {patient.get_name()} | Age: {patient.age} | Gender: {patient.gender}")

    if not patient.visits:
        print("No visits found.")
        return

    for idx, visit in enumerate(patient.visits, start=1):
        print("\n" + "=" * 60)
        print(
            f"Visit {idx}: {visit.visit_id} | Date: {visit.created_on} | "
            f"Dept: {visit.department} | Doctor: {visit.doctor_id}"
        )
        print(f"Diagnosis: {visit.diagnosis}")
        print(f"Presenting Complaints: {visit.clinical_history.presenting_complaints}")

        # ---------- INVESTIGATIONS ----------
        print("Investigations:")
        if not visit.investigations:
            print("  None")
        else:
            for inv in visit.investigations:
                # For free-text investigations loaded from Excel -> inv.inv_type == "Text"
                # Show only the text, not "Text - ..."
                if inv.inv_type.lower() == "text":
                    print(f"  - {inv.name}")
                else:
                    print(f"  - {inv.inv_type} - {inv.name}")

        # ---------- MEDICATIONS ----------
        print("Medications:")
        if not visit.medications:
            print("  None")
        else:
            for med in visit.medications:
                # Excel-loaded meds: Medication("Prescription", full_text, "", "", "")
                if med.name == "Prescription" and med.dose:
                    print(f"  - {med.dose}")
                else:
                    # Build a clean string without trailing 'for .' noise
                    parts = []

                    # Name + dose
                    if med.name:
                        parts.append(med.name.strip())
                    if med.dose:
                        # keep dose attached to name if both exist
                        if parts:
                            parts[-1] = (parts[-1] + " " + med.dose.strip()).strip()
                        else:
                            parts.append(med.dose.strip())

                    # Frequency
                    if med.frequency:
                        parts.append(med.frequency.strip())

                    # Duration -> prefix with 'for ' only if present
                    if med.duration:
                        parts.append(f"for {med.duration.strip()}")

                    # Join main medication string
                    line = ", ".join(parts).strip()

                    # Instructions -> add as a separate sentence only if present
                    if med.instructions:
                        line = (line + f". {med.instructions.strip()}").strip()

                    print("  - " + line)

        print(f"Advice: {visit.advice}")
        print(f"Disposition: {visit.disposition or 'Not recorded'}")


# -------------------------
# Main program
# -------------------------

def main():
    system = HMISOPDSystem()
    system.load_dummy_data_from_excel(EXCEL_FILE)
    while True:
        print("\n========== HMIS OPD MENU ==========")
        print("1. Patient Registration (with visit creation)")
        print("2. Doctor Consultation (complete pending visit)")
        print("3. View Patient History")
        print("4. Exit")
        choice = input("Select option: ").strip()
        if choice == "1":
            registration_flow(system)
        elif choice == "2":
            doctor_consultation_flow(system)
        elif choice == "3":
            view_history_flow(system)
        elif choice == "4":
            print("Exiting HMIS OPD System.")
            break
        else:
            print("Invalid choice. Try again.")

if __name__ == "__main__":

    main()
