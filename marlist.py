def add_list(list):
    sum=0
    for i  in range(0,len(list)):
        sum += list[i]
    return sum    
def get_mark(name,subject):
    return float(input("Enter "+ name +"'s" + subject + "mark: "))
try:
    import xlsxwriter

    noOfStudents = int(input("Enter the number of students :"))
    grade=input("Enter the Class :")
    MAXMARKS=int(input("Enter toltal marks: "))
    noOfsubjects=int(input("Enter the number of subjects: "))
    subjects=[]
    marks=[]
    ExamName = input("Enter the exam name:")
    print("Enter the subject names in order ")
    for i in range(0,noOfsubjects):
        subjects.append()

    for i in range(0,noOfStudents):
        name = input("name of student"+str(i)+": ")
        workbook = xlsxwriter.Workbook(name+'.xlsx')
        ws = workbook.add_worksheet()
        for sub in subjects:
            mark = get_mark(name,sub)
            while mark > MAXMARKS:
                mark = get_mark(name,sub)
            marks.append(mark)    


        merge_format = workbook.add_format({
        'bold':     True,
        'align':    'center',
        'valign':   'vcenter',})
        color_merge = workbook.add_format({
        'bold':     True,
        'align':    'center',
        'valign':   'vcenter',
        'fg_color': '#E57283',})
        blue_merge = workbook.add_format({
            'fg_color': '#bdd7ee',
            'valign':   'vcenter',
        })
        green_merge = workbook.add_format({
            'fg_color': 'c6e0b4',
            'align':    'center',
            'valign':   'vcenter',
        })
        yellow_merge = workbook.add_format({
            'fg_color': '#fff2cc',
            'align': 'center',
            'valign': 'center',
            'font_color':'#203764'
        })
        ws.merge_range('B1:C1','',merge_format)
        ws.merge_range('B2:C2','MAHARISHI VIDYA MANDIR SR. SEC. SCHOOL',color_merge)
        ws.merge_range('B3:C3','Ingur, Erode - 52',color_merge)
        ws.merge_range('B4:C4','',color_merge)
        ws.merge_range('B6:C6',ExamName,yellow_merge)
        ws.merge_range('B8:C8','NAME:'+name,blue_merge)
        ws.merge_range('B9:C9','GRADE:'+ grade,blue_merge)
        ws.merge_range('B11:B12','SUBJECT',green_merge)
        ws.write('C11','(18.12.2020 to 24.12.2020)',green_merge)
        ws.write('C12','Max Marks:'+str(MAXMARKS),green_merge)
        val = 13
        var = 0
        for sub in subjects:
            ws.write('B'+str(val),sub,yellow_merge)
            ws.write('C'+str(val),marks[var],workbook.add_format({
                'align':'center'
            }))
            val+=1
            var+=1
        
        ws.write('B'+str(val+1),'Total',yellow_merge)
        ws.write('C'+str(val+1),add_list(marks))
        ws.write('B'+str(val + 2),'Percentage',yellow_merge)
        ws.write('C'+str(val +2),round(sum/(MAXMARKS*noOfsubjects)*100,2))
        workbook.close()   
except KeyboardInterrupt:
    print("\nClosing")
