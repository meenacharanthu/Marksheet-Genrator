
import PySimpleGUI as sg
import csv
import pandas as pd 
import yagmail
import smtpd
import openpyxl
import os


sg.theme('DarkAmber')

def generate_marksheet():
  def highlight_col (x) :
    r = 'color: blue'
    df1 = pd. DataFrame (' ', index=x. index, columns=x. columns)
    df1. iloc [ :, 1] = r
    return df1  
  correct_ans_data ={}
  student_ans_data ={}
  student_ans_list = []
  mark_list = []
  right_list = []
  wrong_list =[]
  not_attempt_list =[]
  n_right = []
  n_wrong = []
  file_responses = open( resp, 'r')
  reader2 = csv.reader(file_responses)                                # read data with reader.
  next(reader2)
  
  for row in reader2:
      if 'ANSWER' in row[6]:
          correct_ans_data = row[7:35]
      student_ans_data = row[7:35] 
      r,w,t,n =0,0,0,0
      for (a,b) in zip(student_ans_data,correct_ans_data):
         if b in a:
           r = r+1
         elif a in '':
           n = n+1
         else:
           w = w+1
      nr = pos*r
      nw = neg*w
      t = (pos*r)+(neg*w)
      o = (pos*28)
      mark_list.append( str(t)+'/'+str(o))

      right_list.append(r)
      wrong_list.append(w)
      n_right.append(nr)
      n_wrong.append(nw)
      not_attempt_list.append(n)

  movies_df = pd.read_csv(resp)

  df = movies_df[['Name','Roll Number']]
  df2 = movies_df.iloc[:, 7:35]

  d = {'Right(+5)': right_list,'no_rig': n_right,'Wrong(-1)': wrong_list,'no_wro':n_wrong,'Not Attempt(0)': not_attempt_list, 'Total':mark_list}
  df3 = pd.DataFrame(d)
  # print(df3)
  df1= df.join(df3)
  # print(df1)
 
  responses_df = pd.read_csv( resp ,index_col = 'Roll Number')

  for (row,row1,row2,row3) in zip(df2.iterrows(),df.iterrows(),df3.iterrows(),responses_df.iterrows()):
        filepath = 'output/marksheet/' + row3[0] + '.xlsx'
        directory = os.path.dirname(filepath)

        if not os.path.exists(directory):
          os.makedirs(directory)

        data = {'Student Ans':row[1],'Correct Ans': correct_ans_data}
        df4 = pd.DataFrame(data)
        # print(df4)
        df_details = pd.DataFrame(row1[1])
        # print(df_details)
        h = pd.DataFrame(row2[1])
        # print(h)
        No ={'Right':h.iloc[0],'Wrong':h.iloc[2],'Not Attempt':h.iloc[4], 'Max':28}
        No_df = pd.DataFrame(No)
        Marking ={'Right': pos ,'Wrong': neg,'Not Attempt':[0], 'Max':['']}
        Marking_df = pd.DataFrame(Marking)
        Total = {'Right':h.iloc[1],'Wrong':h.iloc[3],'Not Attempt':[''], 'Max':h.iloc[5]}
        Total_df = pd.DataFrame(Total)
        frames = [No_df,Marking_df,Total_df]
        res = pd.concat(frames)
        with pd.ExcelWriter( filepath) as writer:  
           df_details.to_excel(writer, sheet_name='quiz',startrow=5 , startcol=0, header=False, index= True)
           res.to_excel(writer, sheet_name='quiz',startrow= 8 , startcol=0, header=True, index= True)
           df4.style.set_properties(**{'text-align': 'center','border-color':'Black','border-width':'thin','border-style':'solid'}).apply(lambda x: ["color: green; text-align:center; " if x.iloc[0] == x.iloc[1]  else( "color: red; text-align:center;" if v!= x.iloc[1] else "color: blue; text-align:center;") for v in x], axis = 1).apply(highlight_col, axis= None).to_excel(writer,sheet_name='quiz',startrow= 15 , startcol=0, header=True, index= False, engine='openpyxl')
           workbook  = writer.book
           cell_Format = workbook.add_format({'bold': True,'left': 1, 'right': 1, 'top': 1, 'bottom': 1})
           worksheet = writer.sheets['quiz']
           worksheet.insert_image('A1', 'iitp.jpeg')
           worksheet.write('A10','No.',cell_Format)
           worksheet.write('A11','Marking',cell_Format)
           worksheet.write('A12','Total',cell_Format)
           worksheet.write('D6','Exam:',cell_Format)
           worksheet.write('E6','quiz')
def generate_consise_marksheet():
    correct_ans_data ={}
    student_ans_data ={}
    mark_list = []
    right_list = []
    file_responses = open("responses.csv", 'r')
    reader2 = csv.reader(file_responses)                                # read data with reader.
    next(reader2)

    for row in reader2:
       if 'ANSWER' in row[6]:
         correct_ans_data = row[7:35]
       student_ans_data = row[7:35] 
       r,w,t,n =0,0,0,0
       for (a,b) in zip(student_ans_data,correct_ans_data):
           if b in a:
             r = r+1
           elif a in '':
             n = n+1
           else:
             w = w+1
       nr = pos*r
       nw = neg*w
       t = (pos*r)+ (neg*w)
       o = (pos*28)
       data ={'statusAns':'['+str(r)+','+str(w)+','+str(n)+']'}
       marks_data={'Score_After_Negative':str(t) +'/'+str(o)}
       mark_list.append( marks_data)
       right_list.append(data)

    movies_df = pd.read_csv("responses.csv")

    movies_df.rename(columns={
        'Score': 'Google_Score' 
            }, inplace=True)

    df = movies_df.iloc[:,0:6]
    # print(df)
    df1 = pd.DataFrame(mark_list)
    # print(df1)
    df = df.join(df1)
    # print(df)
    df3 = movies_df.iloc[:,6:35]
    df = df.join(df3)
    # print(df)
    df4 = pd.DataFrame(right_list)
    # print(df4)
    df = df.join(df4)
    # print(df)
    df.to_csv('concise_marksheet.csv')

def generate_mail():
  filepath =  "output/marksheet"
  directory = os.path.dirname(filepath)



  k = yagmail.SMTP('meenacharanthu02@gmail.com','Meena@2002')

  file_responses = open("responses.csv", 'r')
  reader2 = csv.reader(file_responses)                                # read data with reader.
  next(reader2)

  for i in reader2:
    roll = i[6]
    k.send(i[1],attachments=filepath+'/'+roll+'.xlsx')
    k.send(i[4],attachments= filepath+'/'+roll+'.xlsx')
    print("email sent to " + roll)
      


layout = [[sg.Text('Gui for project1 window')],      
                 [sg.Text('Browse for Master roll csv')],
                 [sg.Input(key='-MA-'), sg.FileBrowse()],
                 [sg.Text('Browse for Response csv')],
                 [sg.Input(key='-RE-'), sg.FileBrowse()],
                 [sg.Text('Enter the marks for correct and wrong answers')],
                 [sg.Text('Marks for correct answers'), sg.Input(key='-IN-')],
                 [sg.Text('Marks for wrong answers'), sg.Input(key='-ID-')],
                 [sg.Button('Generate Roll no wise marksheet')],
                 [sg.Button('Generate consise marksheet')],
                 [sg.Button('Generate Mail')],
                 [sg.Submit(), sg.Cancel()]]    


window = sg.Window('ORIGINAL').Layout(layout)    
while True:             # Event Loop
    event, values = window.Read()
    ma_roll = values['-MA-']
    resp = values['-RE-']
    pos = int(values['-IN-']) 
    neg = int(values['-ID-'])
    if event in (None, 'Exit'):
        break
    if event == 'Generate Roll no wise marksheet':
        generate_marksheet()
    elif event == 'Generate consise marksheet':
        generate_consise_marksheet()   
    elif event == 'Generate Mail':
        generate_mail()
window.close()
 

