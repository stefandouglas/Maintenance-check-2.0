import pandas as pd
from datetime import datetime
from flask import Flask, request, jsonify

# ✅ File paths (relative to your project directory on Render)
CONVERSATION_FILE = "conversation_tracker.xlsx"
INDUCTION_FILE = "induction_tracker.xlsx"
MAINTENANCE_FILE = "maintenance_schedule.xlsx"

app = Flask(__name__)

###################################
#       CONVERSATION TRACKER      #
###################################

def find_conversation(email, subject):
    try:
        df = pd.read_excel(CONVERSATION_FILE)
        email = str(email).strip().lower()
        subject = str(subject).strip().lower()
        existing_conversation = df[
            (df['Email ID'].str.strip().str.lower() == email) &
            (df['Subject'].str.strip().str.lower() == subject)
        ]
        if not existing_conversation.empty:
            return df, existing_conversation.index[0]
        return df, None
    except Exception as e:
        print(f"❌ Error in find_conversation: {str(e)}")
        return None, None

def determine_next_status(current_status, attachment_present, engineer_names_present):
    if current_status.lower() == 'scheduling request':
        return 'Awaiting RAMS and Engineer Names'
    elif current_status.lower() == 'awaiting rams and engineer names':
        if attachment_present and engineer_names_present:
            return 'Conversation Complete'
        elif attachment_present:
            return 'Awaiting Engineer Names'
        elif engineer_names_present:
            return 'Awaiting RAMS'
    elif current_status.lower() == 'awaiting rams' and attachment_present:
        return 'Conversation Complete'
    elif current_status.lower() == 'awaiting engineer names' and engineer_names_present:
        return 'Conversation Complete'
    return current_status

def get_next_step_instruction(status):
    status = status.lower()
    if status == "scheduling request":
        return "Please ask the contractor to provide a proposed maintenance date."
    elif status == "awaiting rams and engineer names":
        return "Please ask the contractor to provide RAMS and the names of the attending engineers."
    elif status == "awaiting engineer names":
        return "Please ask for the names of the attending engineers."
    elif status == "awaiting rams":
        return "Please ask for RAMS."
    elif status == "conversation complete":
        return "All information has been received. Confirm attendance and say thank you."
    else:
        return "Continue monitoring this conversation."

def update_conversation_status(df, index, new_status):
    try:
        df.at[index, 'Status'] = new_status
        df.at[index, 'Last Updated'] = datetime.now().strftime('%Y-%m-%d')
        df.to_excel(CONVERSATION_FILE, index=False)
        return {"status": "success", "message": f"Conversation updated to {new_status}.", "new_status": new_status}
    except Exception as e:
        print(f"❌ Error in update_conversation_status: {str(e)}")
        return {"status": "error", "message": f"Error updating conversation: {str(e)}"}

def determine_initial_status(attachment_present, engineer_names_present):
    if attachment_present and engineer_names_present:
        return "Conversation Complete"
    elif attachment_present:
        return "Awaiting Engineer Names"
    elif engineer_names_present:
        return "Awaiting RAMS"
    else:
        return "Scheduling Request"

def create_new_conversation(email, subject, initial_status):
    try:
        df = pd.read_excel(CONVERSATION_FILE)
        domain = str(email.split('@')[1]) if '@' in email else "unknown"
        company = domain.split('.')[0] if '.' in domain else domain
        new_data = {
            'Email ID': email,
            'Sender Domain': domain,
            'Company Name': company,
            'Subject': subject,
            'Status': initial_status,
            'Last Updated': datetime.now().strftime('%Y-%m-%d'),
            'Sender Domain + Subject': f"{domain} {subject}"
        }
        df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
        df.to_excel(CONVERSATION_FILE, index=False)
        return {"status": "success", "message": "New conversation created.", "new_status": initial_status}
    except Exception as e:
        print(f"❌ Error in create_new_conversation: {str(e)}")
        return {"status": "error", "message": f"Error creating new conversation: {str(e)}"}

@app.route('/check_conversation', methods=['POST'])
def check_conversation():
    try:
        data = request.get_json()
        print("✅ Incoming data from Zapier:", data)
        email = str(data.get('email', '')).strip()
        subject = str(data.get('email_subject', '')).strip()
        attachment = str(data.get('attachment', 'No')).strip().lower() == 'yes'
        engineers_raw = str(data.get('engineer_names', '')).strip()

        if engineers_raw.lower() in ['none', '', 'null']:
            engineers_raw = ''

        engineers = [e.strip() for e in engineers_raw.split(',')] if engineers_raw else []
        engineer_names_present = bool(engineers)

        if not email or not subject:
            return jsonify({"status": "error", "message": "Missing required fields: email or subject."})

        df, index = find_conversation(email, subject)

        if df is not None and index is not None:
            current_status = df.at[index, 'Status']
            new_status = determine_next_status(current_status, attachment, engineer_names_present)
            instruction = get_next_step_instruction(new_status)
            response_data = update_conversation_status(df, index, new_status)
            response_data["next_step_instruction"] = instruction
            return jsonify(response_data)
        elif df is not None:
            initial_status = determine_initial_status(attachment, engineer_names_present)
            instruction = get_next_step_instruction(initial_status)
            response_data = create_new_conversation(email, subject, initial_status)
            response_data["next_step_instruction"] = instruction
            return jsonify(response_data)
        else:
            return jsonify({"status": "error", "message": "Failed to read conversation tracker."})
    except Exception as e:
        print(f"❌ Error in check_conversation route: {str(e)}")
        return jsonify({"status": "error", "message": f"Error: {str(e)}"})

###################################
#       INDUCTION CHECKER        #
###################################

@app.route('/check_inductions', methods=['POST'])
def check_inductions():
    try:
        data = request.get_json()
        company = data.get("company")
        engineers = data.get("engineers", [])
        maintenance_date_str = data.get("maintenance_date")

        if not maintenance_date_str:
            return jsonify({"status": "error", "message": "Missing maintenance date."})

        df = pd.read_excel(INDUCTION_FILE)
        maintenance_date = datetime.strptime(maintenance_date_str, "%Y-%m-%d").date()

        responses = []
        for engineer in engineers:
            match = df[
                (df['Company'].str.strip().str.lower() == company.strip().lower()) &
                (df['Name'].str.strip().str.lower() == engineer.strip().lower())
            ]
            if match.empty:
                responses.append(f"{engineer} requires an induction.")
            else:
                try:
                    expiry_date = pd.to_datetime(match.iloc[0]['Expiry Date (Auto)']).date()
                    if expiry_date >= maintenance_date:
                        responses.append(f"{engineer} is inducted for the scheduled date.")
                    else:
                        responses.append(f"{engineer}'s induction expires before the scheduled date and needs to be redone.")
                except:
                    responses.append(f"Could not parse expiry date for {engineer}.")

        return jsonify({"status": "success", "results": responses})
    except Exception as e:
        return jsonify({"status": "error", "message": f"Server error: {str(e)}"})

###################################
#      MAINTENANCE CHECKER       #
###################################

def check_maintenance_window(requested_date, equipment_name, company_name):
    df = pd.read_excel(MAINTENANCE_FILE)
    equipment_name = equipment_name.strip().lower()
    company_name = company_name.strip().lower()

    df["Normalized Equipment"] = df["Maintenance subject"].astype(str).str.strip().str.lower()
    df["Normalized Company"] = df["Company"].astype(str).str.strip().str.lower()

    equipment_data = df[(df["Normalized Equipment"] == equipment_name) & (df["Normalized Company"] == company_name)]

    if equipment_data.empty:
        return {"status": "error", "message": f"(❌ No maintenance record found for '{equipment_name}' under '{company_name}'.)"}

    scheduled_months = []
    for quarter in ["Inspection date Q1", "Inspection date Q2", "Inspection date Q3", "Inspection date Q4"]:
        if pd.notna(equipment_data[quarter].values[0]):
            month_name = str(equipment_data[quarter].values[0]).strip()
            scheduled_months.append(month_name)

    requested_month = requested_date.strftime('%B')
    if requested_month in scheduled_months:
        return {"status": "Yes", "message": f"(✅ The requested date {requested_date.date()} is within the maintenance window.)"}
    else:
        next_due_month = scheduled_months[0] if scheduled_months else "Unknown"
        return {"status": f"No - Due in {next_due_month}", "message": f"(❌ The requested date {requested_date.date()} is NOT within the maintenance window. Next due: {next_due_month}.)"}

@app.route('/check_maintenance', methods=['POST'])
def check_maintenance_route():
    try:
        data = request.get_json()
        equipment_name = data.get('equipment_name')
        requested_date = data.get('requested_date')
        company_name = data.get('company_name')

        requested_date = datetime.strptime(requested_date, "%d/%m/%y")
        maintenance_check_result = check_maintenance_window(requested_date, equipment_name, company_name)

        return jsonify({
            "status": "Success",
            "message": "Maintenance check completed successfully.",
            "maintenance_check_result": maintenance_check_result
        })
    except Exception as e:
        return jsonify({"status": "error", "message": f"An error occurred: {str(e)}"})

###################################
#         RUN FLASK APP          #
###################################

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=10000)
