from operator import index
from pptx import Presentation
from pptx.util import Inches
from pptx.util import Pt

###########
# ingest payload
###########

slack_payload={
    "body": {
        "customer": "Awesome Customer",
        "setSA": "First_name Last_Name",
        "useCases": [
            "Cloud Migrations - Data Center Wide",
            "Cloud Migrations - Infrastructure Refresh",
            "Disaster Recovery - New DR"
        ],
        "tvType": "POC",
        "services": [
            "VMware Cloud Disaster Recovery",
            "vRealize Operations Manager"
        ],
        "estDate": "2021-11-04"
    },
    "items": [
        {
            "TestID": "2001",
            "Test Case ": "Connect protected site to Vmware Cloud DR Service",
            "Solutions/add ons": "VCDR",
            "Task Ownership": "VMware & CUSTOMER",
            "Results ": "Protected site appears in VCDR service as \"Protected Site\"",
            "Requirement Mapping": "DRaaS Connector successfully deployed and connected to VMware Cloud DR\nProtected Site vCenter registered with VCDR SaaS orchestrator"
        },
        {
            "TestID": "2002",
            "Test Case ": "Provision cloud SDDC",
            "Solutions/add ons": "VMware Cloud",
            "Task Ownership": "VMware & CUSTOMER",
            "Results ": "Successful instantiation / creation of VMware Cloud on AWS Software Defined Datacenter (SDDC);  if using VCDR, provisioned through VCDR interface\n",
            "Requirement Mapping": "Vmware Account activated\nAWS Account activated\nAWS VPC created with necessary CIDR and subnets\nSDDC created and connected to AWS VPC"
        },
        {
            "TestID": "2003",
            "Test Case ": "Basic connectivity",
            "Solutions/add ons": "VMware Cloud",
            "Task Ownership": "VMware & CUSTOMER",
            "Results ": "Communications between VMC/On-Prem – Direct Connect or IPSec VPN",
            "Requirement Mapping": "L3 VPN successfully established – OR - Direct Connect successfully established\nOpen / interact with vCenter in VMC \nCold migration of VM is successful"
        },
        {
            "TestID": "2008",
            "Test Case ": "Protect VMs for test using VMware Cloud DR solution",
            "Solutions/add ons": "Site Recovery, VCDR",
            "Task Ownership": "VMware & CUSTOMER",
            "Results ": "Virtual machines replicated to Scale-out filesystem",
            "Requirement Mapping": "VMs successfully backed up to VMware Cloud DR\nProtection Groups created\nInput VM list required"
        },
        {
            "TestID": "2009",
            "Test Case ": "Successful DR test",
            "Solutions/add ons": "Site Recovery, VCDR",
            "Task Ownership": "VMware & CUSTOMER",
            "Results ": "Validated virtual machine(s) running successfully in isolated network bubble in DR location",
            "Requirement Mapping": "Configuration of 1-3 recovery plans\nConfiguration of test network for DR\nSuccessful execution of DR test"
        },
        {
            "TestID": "2010",
            "Test Case ": "Successful DR execution",
            "Solutions/add ons": "Site Recovery, VCDR",
            "Task Ownership": "VMware & CUSTOMER",
            "Results ": "Successful failover of 1-10 virtual machines to VMC on AWS",
            "Requirement Mapping": "Configuration of 1-3 recovery plans\nSuccessful execution of DR failover\nSuccessful execution of DR fail back"
        },
        {
            "TestID": "2011",
            "Test Case ": "VMware SDDC Operations from VMware Cloud DR Orchestrator",
            "Solutions/add ons": "VCDR",
            "Task Ownership": "VMware & CUSTOMER",
            "Results ": "VCDR SaaS Orchestrator able to operate / scale SDDC as necessary",
            "Requirement Mapping": "Expand SDDC \nContract SDDC\nAdd network to SDDC\nCreate Firewall Rule\nDeploy VM into SDDC\nCreate resource pools / folders "
        }
    ]
}

###########
# instantiate variables for data ingest
###########
customer = slack_payload["body"]["customer"]
setSA = slack_payload["body"]["setSA"]
tvType = slack_payload["body"]["tvType"]
estDate = slack_payload["body"]["estDate"]
services = slack_payload["body"]["services"]
use_cases = slack_payload["body"]["useCases"]
success_count = len(slack_payload["items"])

###########
# open sample presentation
###########
prs = Presentation('sample.pptx')

###########
# instantiate slide layouts
###########
slide_layout_title = prs.slide_layouts[0]
slide_layout_section = prs.slide_layouts[1]
slide_layout_content = prs.slide_layouts[2]
slide_layout_titlesub = prs.slide_layouts[3]
slide_layout_2content = prs.slide_layouts[4]
slide_layout_blank = prs.slide_layouts[5]
slide_layout_thanks = prs.slide_layouts[6]

# to add a slide use sample below
# slide = prs.slides.add_slide(slide_layout_title)

###########
# update title slide
###########
slide_title = prs.slides[0]

# for shape in slide_title.placeholders:
#     print('%d %s' % (shape.placeholder_format.idx, shape.name))

shape_title = slide_title.placeholders[0]
shape_title.text = customer
shape_tvtype = slide_title.placeholders[10]
shape_tvtype.text = tvType
shape_sa = slide_title.placeholders[11]
shape_sa.text = setSA

###########
# update just the facts slide
###########
slide_jtf = prs.slides[1]

shape_content = slide_jtf.placeholders[14]
table = shape_content.table
cell_cust = table.cell(0, 1)
cell_estdate = table.cell(1, 1)
cell_tvtype = table.cell(2, 1)
cell_sa = table.cell(3, 1)
cell_uses = table.cell(4, 1)
cell_solutions = table.cell(5, 1)

cell_cust.text = customer
cell_estdate.text = setSA
cell_tvtype.text = tvType
cell_sa.text = estDate
cell_uses.text = "\n".join(use_cases)
cell_solutions.text = "\n".join(services)

###########
# update success criteria slide
###########
# slide_success = prs.slides.add_slide(slide_layout_content)
slide_success = prs.slides.add_slide(slide_layout_titlesub)
x, y, cx, cy = Inches(0.67), Inches(1.25), Inches(12), Inches(5)
shape_content = slide_success.shapes.add_table(success_count+1, 3, x, y, cx, cy)
table = shape_content.table

# populate table
cell_hdr1 = table.cell(0, 0)
cell_hdr2 = table.cell(0, 1)
cell_hdr3 = table.cell(0, 2)
cell_hdr1.text = "Test"
cell_hdr2.text = "Expected results"
cell_hdr3.text = "Requirement mapping"

# iterate through success criteria, populate table
# reset case_counter
row_counter = 1
for case in slack_payload["items"]:
    case_test = case.get("Test Case ")
    cell_test =  table.cell(row_counter, 0)
    cell_test.text = case_test
    case_results = case.get("Results ")
    cell_results =  table.cell(row_counter, 1)
    cell_results.text = case_results
    case_reqs = case.get("Requirement Mapping")
    cell_reqs =  table.cell(row_counter, 2)
    cell_reqs.text = case_reqs
    row_counter += 1

# force font size to 10
def iter_cells(table):
    for row in table.rows:
        for cell in row.cells:
            yield cell

for cell in iter_cells(table):
    for paragraph in cell.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(10)

###########
# Save presentation
###########
prs.save('tech_val.pptx')
