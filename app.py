from flask import Flask, render_template, request
from flask_wtf import FlaskForm
from wtforms import DateField, SubmitField, BooleanField
from wtforms.validators import DataRequired
from dateutil.parser import parse
from datetime import datetime, timedelta
import subprocess
import os
from TimeSheeter import TimesheetGenerator
from xhtml2pdf import pisa
import yaml

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your_secret_key_here'

# Load client data
with open('clients.yaml', 'r') as file:
    clients_data = yaml.safe_load(file)

class DateForm(FlaskForm):
    start_date = DateField('Start Date', format='%Y-%m-%d', validators=[DataRequired()])
    end_date = DateField('End Date', format='%Y-%m-%d', validators=[DataRequired()])
    week_totals = BooleanField('Include Week Totals', default=True, render_kw={'checked': True})
    submit = SubmitField('Generate Timesheet and Invoice')

@app.route('/', methods=['GET', 'POST'])
def index():
    form = DateForm()
    if form.validate_on_submit():
        start_date = form.start_date.data
        end_date = form.end_date.data
        
        # Format dates for TimesheetGenerator
        start_date_str = start_date.strftime('%d/%m/%Y')
        end_date_str = end_date.strftime('%d/%m/%Y')

        # Generate Timesheet
        generator = TimesheetGenerator()
        timesheets = generator.generate_timesheet(
            parse(start_date_str, dayfirst=True),
            parse(end_date_str, dayfirst=True).replace(hour=23, minute=59, second=59),
            week_totals=form.week_totals.data
        )

        if timesheets:
            timesheet = timesheets[0]
            client_name = timesheet.client_name
            total_hours, remainder = divmod(timesheet.total_duration.total_seconds(), 3600)
            total_minutes = remainder / 3600  # Convert minutes to decimal hours
            total_hours_decimal = total_hours + total_minutes

            # Get client data
            client_data = clients_data['Clients']['Tacx']  # Assuming Tacx for now
            client_reg_name = client_data['registration_name']

            # Get week numbers and dates from the timesheet data
            timesheet_df = timesheet.time_sheet_df
            start_week = min(timesheet_df['Week_nr'])
            end_week = max(timesheet_df['Week_nr'])
            
            # Get first and last day from the timesheet
            first_day = timesheet_df.iloc[0]['Date']  # First row's full date
            last_day = timesheet_df.iloc[-1]['Date']  # Last row's full date
            logo_path = os.path.abspath('templates/logo.jpg')

            invoice_data = {
                'client': client_data,
                'invoice_date': datetime.now().strftime('%d-%m-%Y'),
                'due_date': (datetime.now() + timedelta(days=30)).strftime('%d-%m-%Y'),
                'reference': 'Georges Meinders',
                'items': [{
                    'quantity': f'{total_hours_decimal:.2f} hours',
                    'description': f'Delivered engineering services to {client_reg_name} for week {start_week} up to and including week {end_week} ({first_day} up to and including {last_day}).',
                    'price': '90,00',
                    'total': f'{total_hours_decimal * 90:.2f}',
                    'vat_rate': '0,00'
                }],
                'vat_calculation_text': f'0.00% VAT on € {total_hours_decimal * 90:.2f} = € 0,00',
                'total_amount': f'€ {total_hours_decimal * 90:.2f}',
                'logo_path': logo_path,
            }

            # Generate invoice HTML
            invoice_html = render_template('invoice.html', **invoice_data)

            # Define PDF path
            pdf_path = f'invoice_{datetime.now().strftime("%Y%m%d")}.pdf'

            # Convert HTML to PDF
            try:
                with open(pdf_path, "w+b") as pdf_file:
                    # Convert HTML to PDF
                    pisa_status = pisa.CreatePDF(
                        invoice_html,                # the HTML to convert
                        dest=pdf_file,               # the output file
                        encoding='utf-8'             # encoding of the HTML
                    )
                
                # Check if PDF generation was successful
                if pisa_status.err:
                    return render_template('timesheet.html',
                                       table="Error generating PDF invoice.",
                                       client_name="",
                                       total_hours="",
                                       form=form)

                # Open PDF file
                if os.name == 'nt':  # Windows
                    os.startfile(pdf_path)
                else:  # Linux/Mac
                    subprocess.run(['xdg-open', pdf_path])

            except Exception as e:
                print(f"Error generating PDF: {str(e)}")
                return render_template('timesheet.html',
                                   table=f"Error generating PDF: {str(e)}",
                                   client_name="",
                                   total_hours="",
                                   form=form)

            # Return the timesheet view
            return render_template('timesheet.html',
                               table=timesheet.time_sheet_df.to_html(index=False),
                               client_name=client_name,
                               total_hours=f"{total_hours:.0f} hours and {remainder//60:.0f} minutes",
                               form=form)
        else:
            return render_template('timesheet.html',
                               table="No timesheets generated.",
                               client_name="",
                               total_hours="",
                               form=form)

    return render_template('timesheet.html', form=form)

if __name__ == '__main__':
    app.run(debug=True)