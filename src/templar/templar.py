import clean_data as cd
import word_tools as wt
from docx import Document
from docx.shared import Inches
from pathlib import Path

dir = '/Users/tom/PycharmProjects/Autotemplate/autotemplate/'
file = 'BR17065S5411 aud 27P rev B.xlsm'

path = dir+file

remove_columns = ['Recommended Actions',
                  'Comment',
                  'Link',
                  'Photo']
new_cols = [(' p&s ID',
             'Hazard_ID')]
xlsx_file = cd.clean_xlsx_table(path,
                                sheet='Master List',
                                head=5,
                                remove_col=remove_columns,
                                clean_hdrs=True,
                                drop_empty=True)

df_dict = xlsx_file.to_dict('records')

heading = ['hazard_id']
tbl_1 = ['asset_name',
         'component',
         'defect_type',
         'defect_intensity']
tbl_2 = ['likelihood_of_failure',
         'consequence_of_failure',
         'holcim_risk']
para_1 = ('comment_input')
para_2 = ('recommended_actions_input')
para = [para_1, para_2]

document = Document('template.docx')
# TODO: provide an option to provide a path to the templating document.

for record in df_dict:
    for each in heading:
        wt.insert_paragraph(document,
                            cd.remove_underscore(str(record[each])),
                            cd.remove_underscore(str(each).title()),
                            para_style='Normal',
                            title_style='PS Heading 3')
    wt.insert_paragraph(document, '')
    tbl_1_data = cd.extract_data(record, tbl_1)
    wt.insert_table(document,
                    len(tbl_1),
                    len(tbl_1_data),
                    tbl_1_data,
                    tbl_style='Plain Table 4')
    wt.insert_paragraph(document, '')
    tbl_2_data = cd.extract_data(record, tbl_2)
    wt.insert_table(document,
                    len(tbl_2),
                    len(tbl_2_data),
                    tbl_2_data,
                    tbl_style='Plain Table 4')
    wt.insert_paragraph(document, '')
    for each in para:
        wt.insert_paragraph(document,
                            cd.remove_underscore(str(record[each])),
                            cd.remove_underscore(str(each).title()),
                            para_style='PS Bullet',
                            title_style='PS Heading 4')

    photo_path = Path()
    if record['location'] != 'No Photo':
        photo_path = record['location']
        document.add_picture(str(photo_path), width=Inches(4))

    document.add_page_break()

document.save('converted_file.docx')
