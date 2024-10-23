import pandas as pd
import os
from flask import Flask, request, render_template, redirect, url_for, flash

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Required for flash messages

# Define the paths for the Excel files
excel_file_path = 'members.xlsx'
subscribed_file_path = 'subscribed.xlsx'

# Create a new Excel file if it doesn't exist
if not os.path.exists(excel_file_path):
    pd.DataFrame(columns=['Name', 'Surname', 'ID Number', 'Phone Number', 'Program of Study', 'Department', 'Membership', 'Membership Start Date', 'Membership End Date', 'Member ID']).to_excel(excel_file_path, index=False, engine='openpyxl')

# Create a new subscribed file if it doesn't exist
if not os.path.exists(subscribed_file_path):
    pd.DataFrame(columns=['Name', 'Surname']).to_excel(subscribed_file_path, index=False, engine='openpyxl')

# Function to generate a unique member ID
def generate_member_id():
    if os.path.exists(excel_file_path):
        existing_data = pd.read_excel(excel_file_path, engine='openpyxl')
        if not existing_data.empty:
            last_id = existing_data['Member ID'].max()
            return last_id + 1
    return 1

# Route for the home page
@app.route('/')
def home():
    return render_template('home.html')

# Route for the add member page
@app.route('/add_member', methods=['GET', 'POST'])
def add_member():
    if request.method == 'POST':
        name = request.form['name']
        surname = request.form['surname']
        id_number = request.form['id_number']
        phone_number = request.form['phone_number']
        program_of_study = request.form['program_of_study']
        department = request.form['department']
        membership = request.form['membership']
        membership_start_date = request.form['membership_start_date']
        membership_end_date = request.form['membership_end_date']
        
        # Check for existing member
        existing_data = pd.read_excel(excel_file_path, engine='openpyxl')
        if not existing_data.empty:
            if ((existing_data['Name'] == name) & 
                (existing_data['Surname'] == surname) & 
                (existing_data['ID Number'] == id_number)).any():
                flash('Member with the same Name, Surname, and ID Number already exists!', 'error')
                return redirect(url_for('add_member'))

        member_id = generate_member_id()
        
        # Save member data to Excel
        member_data = pd.DataFrame({
            'Name': [name],
            'Surname': [surname],
            'ID Number': [id_number],
            'Phone Number': [phone_number],
            'Program of Study': [program_of_study],
            'Department': [department],
            'Membership': [membership],
            'Membership Start Date': [membership_start_date],
            'Membership End Date': [membership_end_date],
            'Member ID': [member_id]
        })
        
        if os.path.exists(excel_file_path):
            existing_data = pd.read_excel(excel_file_path, engine='openpyxl')
            updated_data = existing_data._append(member_data, ignore_index=True)
            updated_data.to_excel(excel_file_path, index=False, engine='openpyxl')
        else:
            member_data.to_excel(excel_file_path, index=False, engine='openpyxl')
        
        flash('Member added successfully!', 'success')
        return redirect(url_for('view_members'))
    return render_template('add_member.html')

# Route for the view members page
@app.route('/view_members')
def view_members():
    if os.path.exists(excel_file_path):
        members = pd.read_excel(excel_file_path, engine='openpyxl')
        print(members.to_dict(orient='records'))  # Verify data reading
        
        # Load the subscribed data
        subscribed_data = pd.read_excel(subscribed_file_path, engine='openpyxl')
        
        # Add a "Subscribed" column to the members' data
        members['Subscribed'] = members.apply(lambda row: 'Subscribed' if ((subscribed_data['Name'] == row['Name']) & (subscribed_data['Surname'] == row['Surname'])).any() else 'Not Subscribed', axis=1)
        
        return render_template('view_members.html', members=members.to_dict(orient='records'))
    return render_template('view_members.html', members=[])

if __name__ == '__main__':
    app.run(debug=True)