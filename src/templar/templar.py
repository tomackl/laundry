import clean_data as clean
import word_tools as word
from docx import Document
from docx.shared import Inches
from pathlib import Path

# ==> cli arguments <==
# todo: the path needs to provided to click
dir = '/Users/tom/PycharmProjects/Autotemplate/autotemplate/'
file = 'BR17065S5411 aud 27P rev B.xlsm'
path = dir+file
work_sheet = 'Master List'

# TODO: provide an option to provide a path to the templating template_doc.
template_doc = Document('template.docx')

# ==> data cleanup arguments <==
# todo: provide a way of describing the columns that need to be removed from the spreadsheet
remove_columns = [
    'Recommended Actions',
    'Comment',
    'Link',
    'Photo'
]

# todo: define a way of replacing columns names with new names.
#       the list below isn't actually implemented within the script.
new_cols = [(
    ' p&s ID',
    'Hazard_ID'
)]

# todo: the following information needs to be provided ... somewhere/somehow
#       - path
#       - worksheet
#       - first row of data (head)
#       - which columns are to be removed
#       - are headers to be cleaned?
#       - are empty columns to be dropped.
#       This information should be separated from the general formatting requirements.

xlsx_file = clean.clean_xlsx_table(
                path,
                sheet=work_sheet,
                head=5,
                remove_col=remove_columns,
                clean_hdrs=True,
                drop_empty=True
                )

df_dict = xlsx_file.to_dict('records')

# ==> Formatting specifications <==
# todo: define a way of scheduling the output template_doc. the following needs to be considered.
#       - how to identify headings
#       - how to identify 'runs' of paragraphs/tables
#       - how to assign styles to paragraphs/tables
#       - do we require a page break at the end of each loop?

heading = ['hazard_id']
tbl_1 = ['asset_name',
         'component',
         'defect_type',
         'defect_intensity'
         ]
tbl_2 = ['likelihood_of_failure',
         'consequence_of_failure',
         'holcim_risk'
         ]
para_1 = 'comment_input'
para_2 = 'recommended_actions_input'
para = [para_1, para_2]

# ==> data manipulation & out document assembly <==
for record in df_dict:
    for each in heading:
        word.insert_paragraph(
            template_doc,
            clean.remove_underscore(str(record[each])),
            clean.remove_underscore(str(each).title()),
            para_style='Normal',
            title_style='PS Heading 3'
        )
    word.insert_paragraph(template_doc, '')

    tbl_1_data = clean.extract_data(record, tbl_1)
    word.insert_table(
        template_doc,
        len(tbl_1),
        len(tbl_1_data),
        tbl_1_data,
        tbl_style='Plain Table 4'
    )

    word.insert_paragraph(template_doc,'')

    tbl_2_data = clean.extract_data(record, tbl_2)
    word.insert_table(
        template_doc,
        len(tbl_2),
        len(tbl_2_data),
        tbl_2_data,
        tbl_style='Plain Table 4'
    )
    word.insert_paragraph(template_doc, '')

    for each in para:
        word.insert_paragraph(
            template_doc,
            clean.remove_underscore(str(record[each])),
            clean.remove_underscore(str(each).title()),
            para_style='PS Bullet',
            title_style='PS Heading 4'
        )

    # todo: do we assume that the photos will just do to the end of the page?
    photo_path = Path()
    if record['location'] != 'No Photo':
        photo_path = record['location']
        template_doc.add_picture(
            str(photo_path),
            width=Inches(4)
        )

    template_doc.add_page_break()

template_doc.save('converted_file.docx')
