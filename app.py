from flask import Flask, render_template, request
from flask_wtf import FlaskForm
from wtforms import DateField, SubmitField
from wtforms.validators import DataRequired
from dateutil.parser import parse
import subprocess
import os
from TimeSheeter import TimesheetGenerator  # Import your TimesheetGenerator class

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your_secret_key_here'  # Change this to a secret key

class DateForm(FlaskForm):
    start_date = DateField('Start Date', format='%Y-%m-%d', validators=[DataRequired()])
    end_date = DateField('End Date', format='%Y-%m-%d', validators=[DataRequired()])
    submit = SubmitField('Generate Timesheet')


@app.route('/', methods=['GET', 'POST'])
def index():
    form = DateForm()
    if form.validate_on_submit():
        start_date_str = form.start_date.data.strftime('%d/%m/%Y')
        end_date_str = form.end_date.data.strftime('%d/%m/%Y')

        # Generate Timesheet using your script
        generator = TimesheetGenerator()
        timesheets = generator.generate_timesheet(parse(start_date_str, dayfirst=True),
                                                  parse(end_date_str, dayfirst=True).replace(hour=23, minute=59,
                                                                                             second=59),
                                                  week_totals=False)

        # Assuming you want to display the first timesheet (if multiple)
        if timesheets:
            timesheet_df = timesheets[0].time_sheet_df
            client_name = timesheets[0].client_name
            total_hours, remainder = divmod(timesheets[0].total_duration.total_seconds(), 3600)
            total_minutes = remainder // 60
            total_hours_display = f"{total_hours:.0f} hours and {total_minutes:.0f} minutes"
            return render_template('timesheet.html',
                                   table=timesheet_df.to_html(index=False),
                                   client_name=client_name,
                                   total_hours=total_hours_display,
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