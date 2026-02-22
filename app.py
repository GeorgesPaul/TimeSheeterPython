from flask import Flask, render_template, request
from flask_wtf import FlaskForm
from wtforms import (DateField, SubmitField, BooleanField, SelectMultipleField,
                     SelectField, FloatField, TextAreaField)
from wtforms.validators import DataRequired, Optional
from dateutil.parser import parse
from datetime import datetime, timedelta, date
import subprocess
import os
import io
import json
from markupsafe import Markup, escape
from TimeSheeter import TimesheetGenerator
from xhtml2pdf import pisa
from pypdf import PdfWriter, PdfReader
import yaml

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your_secret_key_here'

# Load client data
with open('clients.yaml', 'r') as file:
    clients_data = yaml.safe_load(file)

# All possible timesheet columns: (DataFrame column name, form field name, fixed PDF width)
# Width is sized to fit header text without wrapping at 9pt Arial; None = Description takes remaining space.
TIMESHEET_COLUMNS = [
    ('Date',          'col_date',          '2.3cm'),
    ('Day_total',     'col_day_total',      '2.0cm'),
    ('Day',           'col_day',            '1.0cm'),
    ('Start_time',    'col_start_time',     '2.3cm'),
    ('End_time',      'col_end_time',       '2.1cm'),
    ('Duration',      'col_duration',       '2.0cm'),
    ('Week_total',    'col_week_total',     '2.2cm'),
    ('Week_nr',       'col_week_nr',        '1.7cm'),
    ('Week_duration', 'col_week_duration',  '2.6cm'),
    ('Description',   'col_description',    None),
]

class DateForm(FlaskForm):
    invoice_date = DateField('Invoice Date', format='%Y-%m-%d',
                             default=date.today, validators=[DataRequired()])
    start_date = DateField('Start Date', format='%Y-%m-%d', validators=[DataRequired()])
    end_date = DateField('End Date', format='%Y-%m-%d', validators=[DataRequired()])
    clients = SelectMultipleField('Clients', choices=[], render_kw={'size': 10}, validators=[DataRequired()])
    week_totals = BooleanField('Include Week Totals', default=True, render_kw={'checked': True})
    append_timesheet = BooleanField('Append Timesheet to Invoice PDF', default=True, render_kw={'checked': True})
    # Timesheet column selection (shown when append_timesheet is checked)
    col_date          = BooleanField('Date',           default=True)
    col_day_total     = BooleanField('Day Total')
    col_day           = BooleanField('Day')
    col_start_time    = BooleanField('Start Time')
    col_end_time      = BooleanField('End Time')
    col_duration      = BooleanField('Duration',       default=True)
    col_week_total    = BooleanField('Week Total')
    col_week_nr       = BooleanField('Week Nr')
    col_week_duration = BooleanField('Week Duration')
    col_description   = BooleanField('Description',    default=True)
    submit = SubmitField('Generate Timesheet and Invoice')
    # Simple invoice fields
    simple_client      = SelectField('Client', choices=[], validators=[Optional()])
    simple_hours       = FloatField('Hours', validators=[Optional()],
                                    render_kw={'type': 'number', 'step': '0.25',
                                               'min': '0', 'placeholder': '0.00'})
    simple_description = TextAreaField('Description', validators=[Optional()],
                                       render_kw={'rows': '5',
                                                  'placeholder': 'Description of services rendered...'})
    simple_submit = SubmitField('Generate Simple Invoice')

@app.route('/', methods=['GET', 'POST'])
def index():
    form = DateForm()

    # Reload clients data dynamically to catch any changes
    with open('clients.yaml', 'r') as file:
        current_clients_data = yaml.safe_load(file)

    client_choices = [(key, data.get('trade_name', key))
                      for key, data in current_clients_data.get('Clients', {}).items()]
    form.clients.choices = client_choices
    form.simple_client.choices = [('', '— select client —')] + client_choices

    client_rates = {
        key: {'hourly_rate': float(data.get('hourly_rate', 90.0)),
              'currency': data.get('currency', '€')}
        for key, data in current_clients_data.get('Clients', {}).items()
    }

    def render(error=None, timesheets_data=None):
        return render_template('timesheet.html', form=form,
                               client_rates=client_rates,
                               error=error, timesheets_data=timesheets_data)

    # --- Simple invoice ---
    if request.method == 'POST' and form.simple_submit.data:
        inv_date     = form.invoice_date.data
        client_key   = form.simple_client.data
        hours        = form.simple_hours.data
        desc_raw     = (form.simple_description.data or '').strip()

        if not inv_date:
            return render(error="Invoice date is required.")
        if not client_key:
            return render(error="Please select a client for the simple invoice.")
        if not hours or hours <= 0:
            return render(error="Please enter a valid number of hours.")

        client_data = current_clients_data['Clients'].get(client_key)
        if not client_data:
            return render(error=f"Client '{client_key}' not found in clients.yaml.")

        hourly_rate  = float(client_data.get('hourly_rate', 90.0))
        currency     = client_data.get('currency', '€')
        total_price  = hours * hourly_rate
        price_str    = f"{hourly_rate:.2f}".replace('.', ',')
        logo_path    = os.path.abspath('templates/logo.jpg')
        # Preserve line breaks in the PDF description
        desc_html    = escape(desc_raw).replace('\n', Markup('<br>'))

        invoice_data = {
            'client': client_data,
            'invoice_date': inv_date.strftime('%d-%m-%Y'),
            'due_date': (inv_date + timedelta(days=30)).strftime('%d-%m-%Y'),
            'reference': 'Georges Meinders',
            'items': [{
                'quantity': f'{hours:.2f} hours',
                'description': desc_html,
                'price': price_str,
                'total': f'{total_price:.2f}',
                'vat_rate': '0,00',
            }],
            'vat_calculation_text': f'0.00% VAT on {currency} {total_price:.2f} = {currency} 0,00',
            'total_amount': f'{currency} {total_price:.2f}',
            'logo_path': logo_path,
        }

        invoice_html = render_template('invoice.html', **invoice_data)
        pdf_path = f'invoice_{client_key}_{datetime.now().strftime("%Y%m%d%H%M%S")}.pdf'

        try:
            with open(pdf_path, 'w+b') as pdf_file:
                status = pisa.CreatePDF(invoice_html, dest=pdf_file, encoding='utf-8')
            if status.err:
                return render(error=f"Error generating PDF for {client_key}.")
            if os.name == 'nt':
                os.startfile(pdf_path)
            else:
                subprocess.run(['xdg-open', pdf_path])
        except Exception as e:
            return render(error=f"Error generating simple invoice: {str(e)}")

        return render()

    # --- Timesheet invoice ---
    if form.validate_on_submit():
        start_date = form.start_date.data
        end_date = form.end_date.data
        selected_clients = form.clients.data
        
        # Format dates for TimesheetGenerator
        start_date_str = start_date.strftime('%d/%m/%Y')
        end_date_str = end_date.strftime('%d/%m/%Y')

        # Generate Timesheet
        try:
            generator = TimesheetGenerator()
            timesheets = generator.generate_timesheet(
                parse(start_date_str, dayfirst=True),
                parse(end_date_str, dayfirst=True).replace(hour=23, minute=59, second=59),
                week_totals=form.week_totals.data,
                selected_clients=selected_clients
            )
        except Exception as e:
            return render(error=f"Error generating timesheet: {str(e)}")

        if timesheets:
            timesheets_data = []

            for timesheet in timesheets:
                client_name = timesheet.client_name
                total_hours, remainder = divmod(timesheet.total_duration.total_seconds(), 3600)
                total_minutes = remainder / 3600  # Convert minutes to decimal hours
                total_hours_decimal = total_hours + total_minutes

                # Get client data
                client_data = current_clients_data['Clients'].get(client_name)
                if not client_data:
                    continue
                client_reg_name = client_data.get('registration_name', client_name)

                # Get week numbers and dates from the timesheet data
                timesheet_df = timesheet.time_sheet_df
                if timesheet_df.empty:
                    continue
                start_week = min(timesheet_df['Week_nr'])
                end_week = max(timesheet_df['Week_nr'])
                
                # Get first and last day from the timesheet
                first_day = timesheet_df.iloc[0]['Date']  # First row's full date
                last_day = timesheet_df.iloc[-1]['Date']  # Last row's full date
                logo_path = os.path.abspath('templates/logo.jpg')

                hourly_rate = float(client_data.get('hourly_rate', 90.0))
                currency = client_data.get('currency', '€')
                total_price = total_hours_decimal * hourly_rate
                price_str = f"{hourly_rate:.2f}".replace('.', ',')

                inv_date = form.invoice_date.data
                invoice_data = {
                    'client': client_data,
                    'invoice_date': inv_date.strftime('%d-%m-%Y'),
                    'due_date': (inv_date + timedelta(days=30)).strftime('%d-%m-%Y'),
                    'reference': 'Georges Meinders',
                    'items': [{
                        'quantity': f'{total_hours_decimal:.2f} hours',
                        'description': f'Delivered engineering services to {client_reg_name} for week {start_week} up to and including week {end_week} ({first_day} up to and including {last_day}).',
                        'price': price_str,
                        'total': f'{total_price:.2f}',
                        'vat_rate': '0,00'
                    }],
                    'vat_calculation_text': f'0.00% VAT on {currency} {total_price:.2f} = {currency} 0,00',
                    'total_amount': f'{currency} {total_price:.2f}',
                    'logo_path': logo_path,
                }

                # Generate invoice HTML
                invoice_html = render_template('invoice.html', **invoice_data)

                # Define PDF path
                pdf_path = f'invoice_{client_name}_{datetime.now().strftime("%Y%m%d%H%M%S")}.pdf'

                try:
                    if form.append_timesheet.data:
                        # Build filtered column list (name, width) based on selected checkboxes
                        selected_col_info = [
                            (col, width) for col, field_name, width in TIMESHEET_COLUMNS
                            if getattr(form, field_name).data and col in timesheet.time_sheet_df.columns
                        ] or [(col, width) for col, _, width in TIMESHEET_COLUMNS
                              if col in timesheet.time_sheet_df.columns]
                        selected_cols = [col for col, _ in selected_col_info]
                        filtered_df = timesheet.time_sheet_df[selected_cols]

                        # Generate invoice PDF in memory (portrait)
                        invoice_buffer = io.BytesIO()
                        status1 = pisa.CreatePDF(invoice_html, dest=invoice_buffer, encoding='utf-8')
                        if status1.err:
                            return render(error=f"Error generating PDF invoice for {client_name}.")

                        # Generate timesheet PDF in memory (landscape)
                        timesheet_pdf_html = render_template('timesheet_pdf.html',
                            client_name=client_name,
                            first_day=first_day,
                            last_day=last_day,
                            timesheet_cols=[{'name': col, 'width': width} for col, width in selected_col_info],
                            timesheet_rows=filtered_df.fillna('').values.tolist(),
                        )
                        timesheet_buffer = io.BytesIO()
                        status2 = pisa.CreatePDF(timesheet_pdf_html, dest=timesheet_buffer, encoding='utf-8')
                        if status2.err:
                            return render(error=f"Error generating timesheet PDF for {client_name}.")

                        # Merge invoice + timesheet PDFs
                        invoice_buffer.seek(0)
                        timesheet_buffer.seek(0)
                        writer = PdfWriter()
                        for reader in [PdfReader(invoice_buffer), PdfReader(timesheet_buffer)]:
                            for page in reader.pages:
                                writer.add_page(page)
                        with open(pdf_path, 'wb') as f:
                            writer.write(f)
                    else:
                        with open(pdf_path, 'w+b') as pdf_file:
                            pisa_status = pisa.CreatePDF(invoice_html, dest=pdf_file, encoding='utf-8')
                        if pisa_status.err:
                            return render(error=f"Error generating PDF invoice for {client_name}.")

                    # Open PDF file
                    if os.name == 'nt':  # Windows
                        os.startfile(pdf_path)
                    else:  # Linux/Mac
                        subprocess.run(['xdg-open', pdf_path])

                except Exception as e:
                    print(f"Error generating PDF for {client_name}: {str(e)}")
                    return render(error=f"Error generating PDF for {client_name}: {str(e)}")

                timesheets_data.append({
                    'client_name': client_name,
                    'total_hours': f"{total_hours:.0f} hours and {remainder//60:.0f} minutes",
                    'table': timesheet.time_sheet_df.to_html(index=False)
                })

            # Return the timesheet view
            return render(timesheets_data=timesheets_data)
        else:
            return render(error="No timesheets generated.")

    return render()

if __name__ == '__main__':
    app.run(debug=True)