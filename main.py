import PyPDF2 as pdf
import re as r
from openpyxl import load_workbook

file1 = open("myfile.txt", "w")

Submittal_Text = ['SUBMITTALS', 'ACTION SUBMITTALS', 'APPROVALS AND SUBMITTALS', 'INFORMATIONAL SUBMITTALS'
    , 'SUBMITTALS FOR REVIEW', 'CLOSEOUT SUBMITTALS', ' SUBMITTALS', ' ACTION SUBMITTALS', ' APPROVALS AND SUBMITTALS'
    , ' INFORMATIONAL SUBMITTALS', ' SUBMITTALS FOR REVIEW', ' CLOSEOUT SUBMITTALS', '  SUBMITTALS', '  ACTION SUBMITTALS'
    , '  APPROVALS AND SUBMITTALS', '  INFORMATIONAL SUBMITTALS', '  SUBMITTALS FOR REVIEW', '  CLOSEOUT SUBMITTALS',
                  ' SUBMITTAL PROCEDURE']

submittals_num =['1.00', '1.01', '1..02', '1.03', '1.04', '1.05', '1.06', '1.07', '1.08', '1.09', '1.10'
    , '1.11', '1.12', '1.13', '1.15', '1.16', '1.1', '1.2', '1.3', '1.4', '1.5', '1.6']

# open the pdf file
reader = pdf.PdfReader("Example_Spec.pdf")

# extract text and do the search
current_page = 0
for page in reader.pages:
    current_page = current_page + 1
    print(f"-----------------------------------YOUR CURRENT PAGE IS: {current_page}---------------------------------------------")
    pg_text_DELETE = page.extract_text()
    pg_text_list_DELETE =pg_text_DELETE.split(" ")
    print(pg_text_DELETE)
    print(pg_text_list_DELETE)
    # print(pg_text)
    #     this loop finding keyword such as "SUBMITTALS"
    for title in Submittal_Text:
        # x=1.0
        # this while loop is to loop over the potiential sections for the submittal
        for s_num in submittals_num:
            pg_text = page.extract_text()
            Submittal_Reqmts = (s_num+""+title)
            res_search = r.search(Submittal_Reqmts, pg_text)
            res_search = str(res_search)
            # obtain string of this when matched: "<re.Match object"
            res_search = res_search[0:16]
            # print(res_search)
            final_text = []

            if res_search == '<re.Match object':
                print(f"SUBMITTAL SECTION FOUND ON: page {current_page} for {Submittal_Reqmts}")
                #stopping at new section if the Submittal section was found
                # if len(s_num) > 3:
                #     value_looking_for = float(s_num)+0.01
                # else:
                #     value_looking_for = float(s_num)+0.1
                text_list = list(pg_text.split(" "))
                # print(type(text_list))
                value_looking_for ="\n" + s_num
                print(f"this is text_list: {text_list}")
                print(f"value looking for: {value_looking_for}")

                index = 0

                try:
                    index = text_list.index(value_looking_for)
                except:
                    index = text_list.index(s_num)

                print(f"index is : {index}")
                print(f"value found is: {text_list[index]}")
                # obtaining final text
                len_text_list= len(text_list)
                # print(f"length of text_list is: {len_text_list}")

                j = 0
                print(f"length of s_num is {len((s_num))}-----------------")
                #stopping at new section if the Submittal section was found
                if len(s_num) > 3:
                    print("length is greater than 3------------------------------------------------------------")
                    value_stop_at = float(s_num)+0.01
                    value_stop_at_edge_case = float(s_num)+0.02
                else:
                    value_stop_at = float(s_num)+0.1
                    print("length is less than 3---------------------------------------")


                # value_stop_at = str(value_stop_at)
                # value_stop_at = float(value_stop_at)

                print(f"value stop at: {value_stop_at}---------------")
                value_stop_at = str(value_stop_at)
                value_stop_at_edge_case = str(value_stop_at_edge_case)
                print(f"value stop at EGDE: {value_stop_at_edge_case}---------------")

                # str_stop = str(stop)
                value_stop_at = "\n"+value_stop_at
                value_stop_at_edge_case = "\n"+value_stop_at_edge_case

                # print(f"index value to stop at {text_list.index(value_stop_at)}--------")
                # print(f"value stop at: {value_stop_at}---------------")
                # print(float(text_list[index]))
                #s print(value_looking_for)
                # if float(text_list[index]) == float(value_looking_for):
                #     print("match")
                # else:
                #     print("not match")
                while float(text_list[index]) == float(value_looking_for):
                    print(f"this is index: {index}")

                    final_text.append(text_list[index])
                    index = index +1

                    while text_list[index] != value_stop_at and text_list[index] != value_stop_at_edge_case and text_list[index] != value_stop_at[1:]:
                        print(f"{index}---------------------------------------")
                        # new index is mainly for tracking value if it reaches the length of the text_list
                        new_index = index+1
                        # if text_list[index] == edge_case_value_stop_at:
                        #     text_list.append(value_stop_at)
                        #     index = text_list.index(value_stop_at)
                        #     # print(f"{new_index} and {len_text_list}")
                        if text_list[index] == '\nPART':
                            text_list.append(value_stop_at)
                            index = text_list.index(value_stop_at)
                        elif new_index < len_text_list:
                            final_text.append((text_list[index]))
                            # print(final_text)
                            index = index +1

                        else:
                            text_list.append(value_stop_at)
                            index = text_list.index(value_stop_at)
                        print(f"{final_text}---------------------------------------------------")
                        # print(index)
                print(f"hello: {text_list}")
                counter = 0
                # max_section_heading should be located early on the list
                # Anything more than 30 the word section maybe found by will not be the heading
                max_index_heading =40
                try:
                    while counter < max_index_heading:
                        if text_list[counter] == '\nSECTION':
                            index_section = text_list.index('\nSECTION')
                            index_section_end = text_list.index('\nPART')
                            print(f"index end: {index_section}------------------1")
                            print(f"index end: {index_section_end}------------------1")
                            print("hello")
                            if index_section_end < max_index_heading:
                                index_section_title = []
                                print(f"type of index section title is: {type(index_section_title)}")
                                print(type(index_section_title))
                                while index_section < index_section_end:
                                    final_index_section_title = index_section_title.append(text_list[index_section])
                                    index_section = index_section + 1
                                counter = len_text_list +1
                            elif text_list.index('PART') > 0:
                                edge_index_section_end = text_list.index('PART')
                                if edge_index_section_end < max_index_heading:
                                    index_section_title = []
                                    print(f"type of index section title is: {type(index_section_title)}")
                                    print(type(index_section_title))
                                    while index_section < edge_index_section_end:
                                        final_index_section_title = index_section_title.append(text_list[index_section])
                                        index_section = index_section + 1
                                    counter = len_text_list +1

                        elif text_list[counter] == 'SECTION':
                            index_section = text_list.index('SECTION')
                            index_section_end = text_list.index('\nPART')
                            index_section_title = []
                            print(f"index end: {index_section}------------------")
                            print(f"index end: {index_section_end}------------------")
                            print(f"type of index section title is: {type(index_section_title)}")
                            if index_section_end < max_index_heading:
                                while index_section < index_section_end:
                                    final_index_section_title = index_section_title.append(text_list[index_section])
                                    index_section = index_section + 1
                                counter = len_text_list +1

                            elif text_list.index('PART') > 0:
                                edge_index_section_end = text_list.index('PART')
                                if edge_index_section_end < max_index_heading:
                                    index_section_title = []
                                    print(f"type of index section title is: {type(index_section_title)}")
                                    print(type(index_section_title))
                                    while index_section < edge_index_section_end:
                                        final_index_section_title = index_section_title.append(text_list[index_section])
                                        index_section = index_section + 1
                                    counter = len_text_list +1

                        elif text_list[counter] == 'Section':
                            index_section = text_list.index('Section')
                            index_section_end = text_list.index('\nPART')
                            print(f"index end: {index_section}------------------")
                            print(f"index end: {index_section_end}------------------")
                            index_section_title = []
                            print(type(index_section_title))
                            if index_section_end < max_index_heading:
                                while index_section < index_section_end:
                                    final_index_section_title = index_section_title.append(text_list[index_section])
                                    index_section = index_section + 1
                                counter = len_text_list +1
                            elif text_list.index('PART') > 0:
                                edge_index_section_end = text_list.index('PART')
                                if edge_index_section_end < max_index_heading:
                                    index_section_title = []
                                    print(f"type of index section title is: {type(index_section_title)}")
                                    print(type(index_section_title))
                                    while index_section < edge_index_section_end:
                                        final_index_section_title = index_section_title.append(text_list[index_section])
                                        index_section = index_section + 1
                                    counter = len_text_list +1

                        elif text_list[counter] == '\nSection':
                            index_section = text_list.index('\nSection')
                            index_section_end = text_list.index('\nPART')
                            print(f"index end: {index_section}------------------")
                            print(f"index end: {index_section_end}------------------")
                            if index_section_end < max_index_heading:
                                index_section_title = []
                                print(f"type of index section title is: {type(index_section_title)}")
                                print(type(index_section_title))
                                while index_section < index_section_end:
                                    final_index_section_title = index_section_title.append(text_list[index_section])
                                    index_section = index_section + 1
                                counter = len_text_list +1
                            elif text_list.index('PART') > 0:
                                edge_index_section_end = text_list.index('PART')
                                if edge_index_section_end < max_index_heading:
                                    index_section_title = []
                                    print(f"type of index section title is: {type(index_section_title)}")
                                    print(type(index_section_title))
                                    while index_section < edge_index_section_end:
                                        final_index_section_title = index_section_title.append(text_list[index_section])
                                        index_section = index_section + 1
                                    counter = len_text_list +1
                        else:
                            index_section_title = ['Check', 'Dwgs']

                        counter = counter + 1
                except:
                    pass

                print("hello")
                print(index_section_title)
                file1.writelines(f"\nThe following Specification is found on page {current_page} of the document. \n")

                final_section_title_str = ' '.join(index_section_title)
                print(f"index section title: {index_section_title}")
                # print(f"index section title: {final_index_section_title}")
                file1.writelines(final_section_title_str)


                # get final text
                final_text_str = ' '.join(final_text)
                # print(final_text_str)
                print(f"final text: {final_text_str}")
                file1.writelines(final_text_str)

file1.close()

# -----
workbook = load_workbook(filename="Requirement_Log.xlsx")
sheet = workbook.active

file = open('myfile.txt')

# read the content of the file opened
content = file.readlines()

#Create header for rows
sheet['A1'] = 'Specification Section'
sheet['B1'] = 'Description of Submittals'
sheet['C1'] = 'Need Submittal Reviewed By (Date)'
sheet['D1'] = 'Need Submittal By Subcontractor (Date)'


excel_i=2
previous_page_found_line =''
previous_spec_line =''

for lines in content:
    # print(f"these are the lines: {lines}")
    # print(lines[0:7])
    if lines == ' ':
        pass
    elif previous_page_found_line == lines:
        print("previous line for spec page found was the same ")
    else:
        if lines[0:44] =='The following Specification is found on page':
            sheet[f'A{excel_i}'] = lines
            previous_page_found_line = lines
        else:
            if previous_spec_line != lines:
                if lines[0:7] == 'SECTION':
                    sheet[f'A{excel_i}'] = lines
                    previous_spec_line = lines
                    # print(f"prev line {previous_spec_line}------------------------------------------1")
                elif lines[0:7] == 'Section':
                    sheet[f'A{excel_i}'] = lines
                    previous_spec_line = lines
                    # print(f"prev line{previous_spec_line}------------------------------------------2")
                else:
                    sheet[f'B{excel_i}'] = lines
                    previous_spec_line = lines
                    # print(f"final line{previous_spec_line}------------------------------------------3")
    excel_i = excel_i + 1

workbook.save(filename="Requirement_Log.xlsx")

