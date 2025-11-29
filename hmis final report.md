# Hospital Information System (HIS) – OPD Module  

**Final Project Report**  

**Team:**  
- Yogesh Singh (25MM60S01)  
- Rick Maity (25MM60003)  
- Sushobhon Ghosh (25MM60A01)  

**Mentor:** Manali Roy  

---

## 1. Introduction  

### 1.1 Problem Context  

Outpatient Departments (OPDs) in many hospitals still depend heavily on paper-based workflows. This leads to:

- Long patient waiting times due to manual registration and file movement  
- Errors and loss of information because of handwritten records  
- Difficulty in generating reports for audits, management, or government agencies  
- Fragmented view of the patient journey, as data is scattered across multiple registers  

The current challenge is to streamline this OPD workflow, ensure accuracy, and provide a single, reliable view of each patient’s OPD history.  

### 1.2 Project Objective  

The project aims to build a **console-based Hospital Information System (HIS) – OPD Module** in Python that:  

- Digitizes the OPD journey from **registration → consultation → diagnosis → disposition**  
- Maintains a **centralized Electronic Health Record (EHR)-like structure** for OPD visits  
- Ensures data quality via strict input validation (UHID, Aadhaar, DOB, etc.)  
- Stores and reads OPD data using an Excel-based dummy dataset  

### 1.3 Scope  

The module currently covers:  

- Patient registration at the front desk with UHID generation / verification  
- Doctor consultation flow, including vitals, history, diagnosis, investigations, and medications  
- Patient history view for all past visits linked to a UHID  
- Reading and appending OPD records to an Excel file (`hmis_dummy_data.xlsx`)  

---

## 2. Data  

### 2.1 Data Source & Format  

The project uses a synthetic **OPD Patient Dataset** stored in an Excel file:

- **File name:** `hmis_dummy_data.xlsx`  
- **Total records:** 100 OPD entries (dummy data)  
- **Purpose:** To simulate a realistic OPD workflow from registration to visit closure. :contentReference[oaicite:1]{index=1}  

The Python program both **reads from** and **appends to** this Excel file to maintain continuity across sessions.

### 2.2 Key Fields  

Each row in the dataset represents one completed OPD visit and contains:

- **Patient & Demographics**
  - `UHID`: 8-digit unique hospital ID  
  - `Name`, `Gender`  
  - `DOB`: date of birth in `DDMMYYYY` format  
  - `Aadhaar`: 12-digit national ID  
  - `Occupation`  
  - `AddressLine1`, `City`, `State`, `Pincode`  
  - `Mobile`  
  - `EmergencyName`, `EmergencyRelation`, `EmergencyPhone`  

- **Clinical & OPD Workflow**
  - `Department`: e.g., General Medicine, Paediatrics, ENT, etc.  
  - `DoctorId`: mapped to predefined doctor list  
  - `PresentingComplaints`: free-text primary complaint  
  - `Investigations`: investigations advised in that visit (text or structured)  
  - `Medications`: prescriptions generated during consultation  
  - `Diagnosis`: final diagnosis  
  - `Advice`: instructions / counselling / follow-up plan  
  - `Disposition`: Follow up / Discharged / Referred  

### 2.3 Data Flow  

- At startup, the system **loads all rows** from `hmis_dummy_data.xlsx` into Python objects (`Patient`, `Visit`).  
- During runtime, new visits are created, completed, and then **appended back** to the same Excel file as new rows.  
- This creates a simple but effective **EHR-style longitudinal record** for OPD visits.

---

## 3. Questions & Answers  

Below are key conceptual and implementation questions, their answers, and relevant code snippets.  
Each question also notes **who primarily led/solved** that part in the project (as per roles in the presentation).

---

### Q1. How is each patient uniquely identified in the HIS OPD module?  

**Solved by:** Yogesh Singh

#### Answer  

Patients are uniquely identified by a **UHID (Unique Hospital Identification Number)**.  

For new patients, UHID is **generated** from:  

- First 4 digits of DOB (`DDMM`)  
- Last 4 digits of Aadhaar (`XXXX`)  

So UHID = `DDMMXXXX` → exactly **8 digits**.  

For patients loaded from Excel, UHID is read directly from the `UHID` column (must be 8 digits).  

#### Key Code Snippet  

```python
# Helper to keep only digits
def clean_digits(value):
    s = str(value)
    return "".join(ch for ch in s if ch.isdigit())

# Generate UHID: first 4 of DOB + last 4 of Aadhaar
def generate_uhid(dob_digits_8, aadhaar_digits_12):
    first_four = dob_digits_8[:4]      # DDMM
    last_four = aadhaar_digits_12[-4:] # last 4 of Aadhaar
    return first_four + last_four      # 8-digit UHID

class Patient:
    def __init__(self, uhid, name, gender, aadhaar_digits, dob, occupation):
        self.__uhid = uhid
        self.__name = name
        self.gender = gender
        self.__aadhaar_digits = aadhaar_digits
        self.__dob = dob
        self.occupation = occupation
        self.visits = []

    def get_uhid(self):
        return self.__uhid

    def get_name(self):
        return self.__name
```

---

### Q2. How does the system validate Aadhaar, DOB, and UHID during registration and lookups?  

**Solved by:** Yogesh Singh

#### Answer  

To maintain data quality, the system performs strict input validation:

**1. Aadhaar:**

-Must be numeric only

-Must be exactly 12 digits

**2. DOB (`DDMMYYYY`):**

-Must be numeric

-Must be exactly 8 digits

-Must represent a valid calendar date

**3. UHID (for existing patients / history view):**

-Must be numeric

-Must be exactly 8 digits

If any rule fails, the system prints an error and re-prompts the user. 

#### Key Code Snippet  

```python
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
```
---

### Q3. How is the dummy OPD dataset loaded from Excel and converted into objects? 

**Solved by:** Rick Maity

#### Answer  

On startup, the system calls `load_dummy_data_from_excel()` which:

1. Reads `hmis_dummy_data.xlsx` into a pandas DataFrame using `dtype=str` to preserve leading zeros.

2. Iterates over each row:

- Cleans and validates **UHID**, **DOB**, **Aadhaar**.

- Creates or retrieves a **Patient** object.

- Creates a **Visit** object linked to the patient.

- Populates presenting complaints, investigations, medications, diagnosis, advice, and disposition.

3. Marks each such visit as **closed** (these are historical visits).

#### Key Code Snippet  

```python
class HMISOPDSystem:
    def load_dummy_data_from_excel(self, excel_path):
        try:
            df = pd.read_excel(excel_path, dtype=str)
        except FileNotFoundError:
            print(f"Warning: Excel file '{excel_path}' not found. Starting with empty system.")
            return

        for _, row in df.iterrows():
            raw_uhid = (row.get("UHID") or "").strip()
            uhid_digits = clean_digits(raw_uhid)
            if len(uhid_digits) != 8:
                continue

            raw_dob = (row.get("DOB") or "").strip()
            dob_digits = clean_digits(raw_dob)
            if len(dob_digits) != 8:
                continue
            try:
                dob = parse_dob_from_digits(dob_digits)
            except Exception:
                continue

            raw_aadhaar = (row.get("Aadhaar") or "").strip()
            aadhaar_digits = clean_digits(raw_aadhaar)
            if len(aadhaar_digits) != 12:
                continue

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

            department = str(row.get("Department", "General Medicine")).strip()
            doctor_id = str(row.get("DoctorId", "")).strip()
            if doctor_id not in [d["id"] for d in DOCTORS]:
                doctor = get_doctor_for_department(department)
                doctor_id = doctor["id"] if doctor else "DR001"

            visit = self.create_visit(uhid, department, doctor_id)

            presenting = str(row.get("PresentingComplaints", "")).strip()
            if presenting:
                history = ClinicalHistory(presenting_complaints=presenting)
                visit.add_clinical_history(history)

            inv_text = str(row.get("Investigations", "")).strip()
            if inv_text:
                visit.add_investigation(Investigation("Text", inv_text))

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
```

---

### Q4. What is the complete front-desk registration and visit creation workflow?

**Solved by:** Rick Maity

#### Answer  

The registration flow (Menu Option 1) performs the following steps:

1. Ask if the patient already has a UHID.

- If **yes**: validate UHID, fetch patient, and create a new visit with selected department and doctor.

- If **no**: capture patient details, validate DOB and Aadhaar, generate UHID, and create both patient and visit.

2. Allow front desk to select the **department** from a predefined list.

3. Auto-assign a doctor based on department using `get_doctor_for_department()`.

4. Create a `Visit` object with status “open” (pending doctor consultation).

#### Key Code Snippet  

```python
def registration_flow(system):
    print("\n--- Patient Registration Desk ---")
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
```

---

### Q5. How is the doctor consultation flow implemented, from vitals to case sheet generation?

**Solved by:** Rick Maity

#### Answer  

The doctor consultation flow (Menu Option 2):

1. Identifies the patient using UHID and chooses an **open visit**.

2. Records **vitals** (`BP`, `Pulse`, `SpO2`, `Temperature`).

3. Captures a detailed **clinical history**: presenting complaints, past history, surgeries, allergies, medications, addictions.

4. Records **diagnosis**.

5. Allows adding multiple **investigations** and **medications**.

6. Captures **advice** and **disposition**, then closes the visit.

7. Generates and prints an **OPD Case Sheet** and appends the visit to Excel.

#### Key Code Snippet  

```python
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

    # Investigations
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

    # Medications
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
```

---

### Q6. How can we view a patient’s complete OPD history and ensure clean display of investigations & medications?

**Solved by:** Sushobhon Ghosh

#### Answer

The history viewer (Menu Option 3):

1. Takes UHID as input and fetches the corresponding patient.

2. Iterates over all visits associated with that patient.

3. Displays visit ID, date, department, doctor, diagnosis, and presenting complaints.

4. For **investigations**, it formats them depending on whether they were loaded as free text or structured.

5. For **medications**, it removes extra words like “text” or “Prescription” and avoids awkward strings like `for .`.

#### Key Code Snippet

```python
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

        # Investigations
        print("Investigations:")
        if not visit.investigations:
            print("  None")
        else:
            for inv in visit.investigations:
                if inv.inv_type.lower() == "text":
                    print(f"  - {inv.name}")
                else:
                    print(f"  - {inv.inv_type} - {inv.name}")

        # Medications
        print("Medications:")
        if not visit.medications:
            print("  None")
        else:
            for med in visit.medications:
                if med.name == "Prescription" and med.dose:
                    print(f"  - {med.dose}")
                else:
                    parts = []
                    if med.name:
                        parts.append(med.name.strip())
                    if med.dose:
                        if parts:
                            parts[-1] = (parts[-1] + " " + med.dose.strip()).strip()
                        else:
                            parts.append(med.dose.strip())
                    if med.frequency:
                        parts.append(med.frequency.strip())
                    if med.duration:
                        parts.append(f"for {med.duration.strip()}")
                    line = ", ".join(parts).strip()
                    if med.instructions:
                        line = (line + f". {med.instructions.strip()}").strip()
                    print("  - " + line)

        print(f"Advice: {visit.advice}")
        print(f"Disposition: {visit.disposition or 'Not recorded'}")
```

---

### Q7. What are the broader implications, benefits, and trade-offs of using an EHR-based OPD module?

**Solved by:** Sushobhon Ghosh

#### Answer

Based on the literature and charts presented:

**Benefits of Electronic Health Records (EHR) / HIS-based OPD:**

1. Improved patient outcomes due to better continuity of care, fewer errors, and more complete information at the point of care.

2. Higher data quality: legible, searchable, and structured data.

3. Faster reporting for internal KPIs, regulatory submissions, and audits.

4. Analytics-ready data enabling future modules like dashboards, disease surveillance, and resource planning.

5. Global trend: Many Western countries have significantly increased EHR adoption in hospitals and clinics, indicating proven value and maturity of such systems.

**Challenges / Trade-offs:**

1. Implementation cost and complexity (hardware, software, training).

2. Change management: staff need time and support to adapt from paper to digital workflows.

3. Data security and privacy: requires proper access control and compliance.

Our project is a simplified console prototype, but it reflects the core ideas of EHR: a **centralized, longitudinal electronic record** of OPD interactions.

---

## 4. References

1. [Python Documentation.](https://docs.python.org/3/)

2. Python Module and Libraries: [datetime](https://docs.python.org/3/library/datetime.html) and [pandas](https://pandas.pydata.org/).
 
3. [Introduction to Computation and Programming Using Python.](https://mitpress.mit.edu/9780262542364/introduction-to-computation-and-programming-using-python/)

4. [Electronic Health Records Statistics 2025 By Healthcare, Data, Management.](https://media.market.us/electronic-health-records-statistics/)

