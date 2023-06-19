import requests
from bs4 import BeautifulSoup
# requests isn't scraping arabic strings, so i have used pandas to get the students name and calculating the student's total result
import pandas as pd
import numpy as np


def Got_to_result_link(url1):
    req1=requests.get(url1).text
    soup1=BeautifulSoup(req1,'html.parser')
    find=soup1.find("tr",{"class":"ewTableRow","onmouseover":"ew_MouseOver(this);"})
    return find

def Enter_to_result_page(find):
    find2=find.find("a").get("href")
    url2=f"http://app1.helwan.edu.eg/EngMatrya/{find2}"
    req2=requests.get(url2).text
    
    soup2=BeautifulSoup(req2,'html.parser')
    return soup2, url2

def estimate_grade(percentage,degrees):
    if percentage>=85:
        degrees.append("م")
    elif percentage>=75:
        degrees.append("ج.ج")
    elif percentage>=65:
        degrees.append("ج")
    elif percentage>=50:
        degrees.append("ل")
    elif percentage>=40:
        degrees.append("ض")
    else:
        degrees.append("ض.ج") 
    return degrees

        
def collceting_current_st_data(soup2,url2,No_of_Subjects,total,current_ip):
    deg=[]
    degrees=[]
    degrees_in_int=[]
    C1_degree=soup2.find_all("td",{"align":"center","width":"100"})
    C2_degree=soup2.find_all("td",{"align":"center","width":"81"})
    for i in C1_degree:
            find=i.find("b")
            if len(find.text)<3 or len(find.text)==3:
                deg.append(find.text)

    
    for i in C2_degree:
        find=i.find("b")
        if len(find.text)<3 or len(find.text)==3 :
            deg.append(find.text)

    for i in range(No_of_Subjects):
        degrees.append(deg[i])
        
    for i in degrees:
            try:
                degrees_in_int.append(int(i))
            except: 
                print("Nan")
    
        
    percentage=""
    try:
        st_degrees=pd.DataFrame(degrees_in_int).sum().loc[0]
        degrees.append(st_degrees)

        percentage=st_degrees*100/total
        degrees.append(f"{round(percentage,2)}%")
    except:
        print("Nan")
        degrees.append("")
        degrees.append("")
    name_df=pd.read_html(url2)
    name=name_df[1].loc[2,1]
    degrees.insert(0,name)
    degrees.insert(0,current_ip)
    degrees = estimate_grade(percentage,degrees)
    print(degrees)
    return degrees




def Finishing(data,C_Names):
    df=pd.DataFrame(data,columns=C_Names)
    df.sort_values(by = "Perc.%", ascending=False, inplace=True)
    df.index=np.arange(1,len(data)+1).tolist()
    gr=df["Grade"].value_counts()
    gr=pd.DataFrame(gr)
    gr.columns=["العدد"]
    print(df)
    df.to_excel("deg.xlsx")
    gr.to_excel("gr.xlsx")


data=[]
C_Names=["ID","Name","Analysis","Heat","Stress","Elec.","Meas.","kin.","Field-1","Total","Perc.%","Grade"]

No_of_Subjects=7
total= 725
first_ip= 25001
final_ip= 25016
current_ip=first_ip
finishing=final_ip+1

while current_ip != finishing:
    try:
        print(current_ip)

        url1=f"http://app1.helwan.edu.eg/EngMatrya/HasasnUpMlist.asp?z_dep=%3D&z_st_name=LIKE&z_st_settingno=%3D&x_st_settingno=&x_st_name=&z_gro=%3D&x_gro=%C7%E1%CB%C7%E4%ED%C9&x_dep=%CA%D5%E3%ED%E3+%E3%ED%DF%C7%E4%ED%DF%ED&z_sec=LIKE&x_sec=&psearch={current_ip}&Submit=++++%C8%CD%CB++++"

        find= Got_to_result_link(url1)
        try:
            soup2,url2 = Enter_to_result_page(find)

            degrees= collceting_current_st_data(soup2,url2, No_of_Subjects, total, current_ip)
            data.append(degrees)
        except:
            print("Not exist")   

        current_ip += 1
    except:
         print("error")
         breaking=input("braking and finishing? \n 1 to break \n 0 to try again")
         if breaking == "1" :
            break
for i in range(3):
    try:     
        Finishing(data,C_Names)
        break
    except:
        print("close the file ")
        try_again=input("to try again \n Enter 1")
