from flask import Flask, render_template, request
from flask_sqlalchemy import SQLAlchemy
# from send_mail import send_mail
from dotenv import load_dotenv
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import *
import gspread
import os
load_dotenv()


# print(os.getenv('TEST'))
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive.file','https://www.googleapis.com/auth/drive']

creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)

client = gspread.authorize(creds)

sheet = client.open('Nest_Test').sheet1

app = Flask(__name__)

ENV = 'dev'

if ENV == 'dev':
    app.debug = True
    app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('POSTGRES')
else:
    app.debug = False
    app.config['SQLALCHEMY_DATABASE_URI'] = ''

app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

class AreaCosts(db.Model):
    __tablename__ = 'area_costs'
    id = db.Column(db.Integer, primary_key=True)
    pjm = db.Column(db.String(40))
    project = db.Column(db.String(4), unique=True)
    address = db.Column(db.String(200))

    building_and_zoning_permit = db.Column(db.Float, default=0)
    septic_permit = db.Column(db.Float, default=0)
    well_permit = db.Column(db.Float, default=0)
    dock_permit = db.Column(db.Float, default=0)
    shoreline_permit = db.Column(db.Float, default=0)
    hoa_app_fee = db.Column(db.Float, default=0)
    hoa_deposit = db.Column(db.Float, default=0)
    surveying = db.Column(db.Float, default=0)

    lot_clearing = db.Column(db.Float, default=0)
    silt_fence = db.Column(db.Float, default=0)
    lot_grading = db.Column(db.Float, default=0)
    demolition = db.Column(db.Float, default=0)
    temp_drive = db.Column(db.Float, default=0)
    culvert_pipe =db.Column(db.Float, default=0)
    soil_compaction_testing = db.Column(db.Float, default=0)
    soil_compaction_eng = db.Column(db.Float, default=0)

    slab_gravel_fill = db.Column(db.Float, default=0)
    foundation_grading = db.Column(db.Float, default=0)
    foundation_overage = db.Column(db.Float, default=0)
    footing_overage =db.Column(db.Float, default=0)
    retaining_wall = db.Column(db.Float, default=0)

    water_tap_fee =db.Column(db.Float, default=0)
    well = db.Column(db.Float, default=0)
    sewer_tap_fee =db.Column(db.Float, default=0)
    septic_system = db.Column(db.Float, default=0)
    grinder_pump = db.Column(db.Float, default=0)

    driveway = db.Column(db.Float, default=0)
    parking_pad = db.Column(db.Float, default=0)
    sidewalks = db.Column(db.Float, default=0)
    additional_flatwork =db.Column(db.Float, default=0)
    pump_truck = db.Column(db.Float, default=0)

    generator_retnal = db.Column(db.Float, default=0)
    fuel_supply = db.Column(db.Float, default=0)

    irrigation = db.Column(db.Float, default=0)
    seed_sod = db.Column(db.Float, default=0)
    shrubs = db.Column(db.Float, default=0)
    final_grading = db.Column(db.Float, default=0)
    pressure_wash = db.Column(db.Float, default=0)
    mailbox = db.Column(db.Float, default=0)
    street_cleaning = db.Column(db.Float, default=0)
    dock = db.Column(db.Float, default=0)
    shoreline = db.Column(db.Float, default=0)

    contingency = db.Column(db.Float, default=0)

    def __init__(self,pjm,project,address,building_and_zoning_permit,septic_permit,well_permit,dock_permit,shoreline_permit,hoa_app_fee,hoa_deposit,surveying,lot_clearing,silt_fence,lot_grading,demolition,temp_drive,culvert_pipe,soil_compaction_testing,soil_compaction_eng,slab_gravel_fill,foundation_grading,foundation_overage,footing_overage,retaining_wall,water_tap_fee,well,sewer_tap_fee,septic_system,grinder_pump,driveway,parking_pad,sidewalks,additional_flatwork,pump_truck,generator_retnal,fuel_supply,irrigation,seed_sod,shrubs,final_grade,pressure_wash,mailbox,street_cleaning,dock,shoreline,contingency):
        self.project = project
        self.pjm = pjm
        self.address = address
        # permitting and setup
        self.building_and_zoning_permit = building_and_zoning_permit
        self.septic_permit = septic_permit
        self.well_permit = well_permit
        self.dock_permit = dock_permit
        self.shoreline_permit = shoreline_permit
        self.hoa_app_fee = hoa_app_fee
        self.hoa_deposit = hoa_deposit
        self.surveying = surveying
        # grading
        self.lot_clearing = lot_clearing
        self.silt_fence = silt_fence
        self.lot_grading = lot_grading
        self.demolition = demolition
        self.temp_drive = temp_drive
        self.culvert_pipe = culvert_pipe
        self.soil_compaction_testing = soil_compaction_testing
        self.soil_compaction_eng = soil_compaction_eng
        # foundation
        self.slab_gravel_fill = slab_gravel_fill
        self.foundation_grading = foundation_grading
        self.foundation_overage = foundation_overage
        self.footing_overage = footing_overage
        self.retaining_wall = retaining_wall
        # sewer and water
        self.water_tap_fee = water_tap_fee
        self.well = well
        self.sewer_tap_fee = sewer_tap_fee
        self.septic_system = septic_system
        self.grinder_pump = grinder_pump
        # concrete
        self.driveway = driveway
        self.parking_pad = parking_pad
        self.sidewalks = sidewalks
        self.additional_flatwork = additional_flatwork
        self.pump_truck = pump_truck
        # framing / mechanical
        self.generator_retnal = generator_retnal
        self.fuel_supply = fuel_supply
        # lanscaping
        self.irrigation = irrigation
        self.seed_sod = seed_sod
        self.shrubs = shrubs
        self.final_grade = final_grade
        self.pressure_wash = pressure_wash
        self.mailbox = mailbox
        self.street_cleaning = street_cleaning
        self.dock = dock
        self.shoreline = shoreline
        # contingency
        self.contingency = contingency

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    if request.method == 'POST':
        # 
        propType = request.form['propType']
        # 
        pjm = request.form['pjm']
        project = request.form['project'].upper()
        address = request.form['address']
        # permitting and setup
        building_and_zoning_permit = request.form['building_and_zoning_permit']
        septic_permit = request.form['septic_permit']
        well_permit = request.form['well_permit']
        dock_permit = request.form['dock_permit']
        shoreline_permit = request.form['shoreline_permit']
        hoa_app_fee = request.form['hoa_app_fee']
        hoa_deposit = request.form['hoa_deposit']
        surveying = request.form['surveying']
        # grading
        lot_clearing = request.form['lot_clearing']
        silt_fence = request.form['silt_fence']
        lot_grading = request.form['lot_grading']
        demolition = request.form['demolition']
        temp_drive = request.form['temp_drive']
        culvert_pipe = request.form['culvert_pipe']
        soil_compaction_testing = request.form['soil_compaction_testing']
        soil_compaction_eng = request.form['soil_compaction_eng']
        # foundation
        slab_gravel_fill = request.form['slab_gravel_fill']
        foundation_grading = request.form['foundation_grading']
        foundation_overage = request.form['foundation_overage']
        footing_overage = request.form['footing_overage']
        retaining_wall = request.form['retaining_wall']
        # sewer and water
        water_tap_fee = request.form['water_tap_fee']
        well = request.form['well']
        sewer_tap_fee = request.form['sewer_tap_fee']
        septic_system = request.form['septic_system']
        grinder_pump = request.form['grinder_pump']
        # conrete
        driveway = request.form['grinder_pump']
        parking_pad = request.form['parking_pad']
        sidewalks = request.form['sidewalks']
        additional_flatwork = request.form['additional_flatwork']
        pump_truck = request.form['pump_truck']
        # framing and mechanical
        generator_retnal = request.form['generator_retnal']
        fuel_supply = request.form['fuel_supply']
        # landscaping
        irrigation = request.form['irrigation']
        seed_sod = request.form['seed_sod']
        shrubs = request.form['shrubs']
        final_grade = request.form['final_grade']
        pressure_wash = request.form['pressure_wash']
        mailbox = request.form['mailbox']
        street_cleaning = request.form['street_cleaning']
        dock = request.form['dock']
        shoreline = request.form['shoreline']
        # contingency
        contingency = request.form['contingency']

        values = [building_and_zoning_permit,septic_permit,well_permit,dock_permit,shoreline_permit,hoa_app_fee,hoa_deposit,surveying,lot_clearing,silt_fence,lot_grading,demolition,temp_drive,culvert_pipe,soil_compaction_testing,soil_compaction_eng,slab_gravel_fill,foundation_grading,foundation_overage,footing_overage,retaining_wall,water_tap_fee,well,sewer_tap_fee,septic_system,grinder_pump,driveway,parking_pad,sidewalks,additional_flatwork,pump_truck,generator_retnal,fuel_supply,irrigation,seed_sod,shrubs,final_grade,pressure_wash,mailbox,street_cleaning,dock,shoreline,contingency]

        start_cell = 9

        if pjm == 'Project Manager' or project == '':
            return render_template('index.html', message=f'Please entire required fields.')

        if db.session.query(AreaCosts).filter(AreaCosts.project == project).count() == 0:
            data = AreaCosts(pjm,project,address,building_and_zoning_permit,septic_permit,well_permit,dock_permit,shoreline_permit,hoa_app_fee,hoa_deposit,surveying,lot_clearing,silt_fence,lot_grading,demolition,temp_drive,culvert_pipe,soil_compaction_testing,soil_compaction_eng,slab_gravel_fill,foundation_grading,foundation_overage,footing_overage,retaining_wall,water_tap_fee,well,sewer_tap_fee,septic_system,grinder_pump,driveway,parking_pad,sidewalks,additional_flatwork,pump_truck,generator_retnal,fuel_supply,irrigation,seed_sod,shrubs,final_grade,pressure_wash,mailbox,street_cleaning,dock,shoreline,contingency)
            db.session.add(data)
            db.session.commit()

            sheet.update('B4',pjm)
            sheet.update('B5',project)
            sheet.update('B6',address)

            for x in range(len(values)):
                sheet.update(f'C{start_cell}', values[x])
                sheet.format(f'C{start_cell}',{
                    'numberFormat':{
                        'type':'CURRENCY',
                        'pattern':'"$" #,##0.00'
                    }
                })
                start_cell += 1
        #     send_mail(pjm,project,address,propType)
            return render_template('success.html')

        return render_template('index.html', message=f'This property already exists.')
if __name__ == '__main__':
    app.run()


# pjm
# project
# address

# building_and_zoning_permit
# septic_permit
# well_permit
# dock_permit
# shoreline_permit
# hoa_app_fee
# hoa_deposit
# surveying

# lot_clearing
# silt_fence
# lot_grading
# demolition
# temp_drive
# culvert_pipe
# soil_compaction_testing
# soil_compaction_eng

# slab_gravel_fill
# foundation_grading
# foundation_overage
# footing_overage
# retaining_wall

# water_tap_fee
# well
# sewer_tap_fee
# septic_system
# grinder_pump

# driveway
# parking_pad
# sidewalks
# additional_flatwork
# pump_truck

# generator_retnal
# fuel_supply

# irrigation
# seed_sod
# shrubs
# final_grade
# pressure_wash
# mailbox
# street_cleaning
# dock
# shoreline

# contingency

# building_and_zoning_permit_comment
# septic_permit_comment
# well_permit_comment
# dock_permit_comment
# shoreline_permit_comment
# hoa_app_fee_comment
# hoa_deposit_comment
# surveying_comment

# slab_gravel_fill_comment
# foundation_grading_comment
# foundation_overage_comment
# footing_overage_comment
# retaining_wall_comment

# lot_clearing_comment
# silt_fence_comment
# lot_grading_comment
# demolition_comment
# temp_drive_comment
# culvert_pipe_comment
# soil_compaction_testing_comment
# soil_compaction_eng_comment

# water_tap_fee_comment
# well_comment
# sewer_tap_fee_comment
# septic_system_comment
# grinder_pump_comment

# driveway_comment
# parking_pad_comment
# sidewalks_comment
# additional_flatwork_comment
# pump_truck_comment

# generator_retnal_comment
# fuel_supply_comment

# irrigation_comment
# seed_sod_comment
# shrubs_comment
# final_grade_comment
# pressure_wash_comment
# mailbox_comment
# street_cleaning_comment
# dock_comment
# shoreline_comment

# print(f"{pjm}\n{project}\n{address}\n{propType}\n{building_and_zoning_permit}\n{septic_permit}\n{well_permit}\n{dock_permit}{shoreline_permit}\n{hoa_app_fee}\n{hoa_deposit}\n{surveying}\n{lot_clearing}\n{silt_fence}\n{lot_grading}\n{demolition}\n{temp_drive}\n{culvert_pipe}\n{soil_compaction_testing}\n{soil_compaction_eng}\n{slab_gravel_fill}\n{foundation_grading}\n{foundation_overage}\n{footing_overage}\n{retaining_wall}\n{water_tap_fee}\n{well}\n{sewer_tap_fee}\n{septic_system}\n{grinder_pump}\n{driveway}\n{parking_pad}\n{sidewalks}\n{additional_flatwork}\n{pump_truck}\n{generator_retnal}\n{fuel_supply}\n{irrigation}\n{seed_sod}\n{shrubs}\n{pressure_wash}\n{mailbox}\n{street_cleaning}\n{dock}\n{shoreline}\n{contingency}")