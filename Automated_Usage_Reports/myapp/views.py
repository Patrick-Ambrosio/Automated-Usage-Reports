from django.shortcuts import render, redirect
from .forms import MyModelForm
from .models import MyModel
import pandas as pd
import os
from datetime import datetime
from datetime import date, timedelta

#Setting the date to the month of the report
prev = date.today().replace(day=1) - timedelta(days=1)
month = prev.strftime("%B")
year = datetime.now().year
current_date = r"%s/1/%s"%(prev.strftime("%m"),year)


# Create your views here.

def postcreate(request):
    if request.method == 'POST':
        form=MyModelForm(request.POST or None,request.FILES or None)
        files = request.FILES.getlist('files')
        provider = request.POST.get("provider")
        user_name = request.POST.get("user_name")
        if form.is_valid():
            
            #add everything you want to add here
            if files:#check if user has uploaded some files
                for f in files:
                    MyModel.objects.create(files=f)
            
            if provider == 'adobe':
                adobe_report(user_name)
            if provider == 'dstillery':
                dstillery_report(user_name)
            if provider == 'eyeota':
                eyeota_report(user_name)
            if provider == 'fyllo':
                fyllo_report(user_name)
            if provider == 'icx':
                icx_report(user_name)  
            if provider == 'neustar':
                neustar_report(user_name)          
            if provider == 'comscoretv':
                comscoretv_report(user_name)
            if provider == 'comscorepa':
                comscorepa_report(user_name)

            return render(request,'success.html')
    else:
        form = MyModelForm()
    return render(request,'main.html',{'form':form})


def fyllo_report(name):
    user = name


    #Importing files. Make sure to add "_Month" at the end of the file

    data = pd.read_excel(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\Fyllo_crosstab_%s.xlsx"%(user,month))
    rx = pd.read_excel(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\Fyllo_RX_crosstab_%s.xlsx"%(user,month))

    #Deletes first row of excel file
    fyllo_dsp = data.drop(data.index[0])
    fyllo_rx = rx.drop(rx.index[0])


    #Calculating net cost
    fyllo_dsp["Gross Cost"]=fyllo_dsp["Total Impressions"]*fyllo_dsp["Segment CPM"]/1000
    fyllo_rx["Gross Cost"]=fyllo_rx["Impressions"]*fyllo_rx["CPM"]/1000    
    gcost = fyllo_dsp.loc[:,"Gross Cost"]
    gcost_rx = fyllo_rx.loc[:,"Gross Cost"]
    ncost_rx = gcost_rx*0.8
    ncost = gcost*0.8
    fyllo_dsp["Net Cost"] = ncost
    fyllo_rx["Net Cost"]=ncost_rx


    #Adding in necessary columns
    fyllo_dsp["Event_date"]=current_date
    fyllo_rx["Event_date"]=current_date
    fyllo_rx["Campaign"]=""
    fyllo_rx["Advertiser"]="PMP - not available"
    fyllo_rx["Agency"]=""

    #renaming columns
    fyllo_dsp.rename(columns = {"Ext. Provider ID": "Segment ID","Segment ":"Segment", "Segment CPM":"Segment Cost", "Campaign Name": "Campaign" }, inplace = True)
    fyllo_rx.rename(columns = {"External Provider ID":"Segment ID","Display Name":"Segment","CPM":"Segment Cost", "Impressions":"Total Impressions"}, inplace =True)

    #Re-ordering the columns to the proper format
    df1 = fyllo_dsp.loc[:,["Event_date","Campaign","Segment ID","Segment","Segment Cost","Advertiser", "Agency","Total Impressions", "Gross Cost", "Net Cost"]]
    rx1 = fyllo_rx.loc[:,["Event_date","Campaign","Segment ID","Segment","Segment Cost","Advertiser", "Agency","Total Impressions", "Gross Cost", "Net Cost"]]

    

    #Renaming the columns
 

    #Deletes any columns with 0 net cost
    df1.drop(df1[df1['Net Cost'] == 0].index, inplace = True)
    rx1.drop(rx1[rx1['Net Cost'] == 0].index, inplace = True)

    #Combines RX and DSP files
    frames = [df1, rx1]
    result = pd.concat(frames)

    #Formating the excel
    result['Segment Cost'] = result['Segment Cost'].map('${:,.2f}'.format)
    result['Gross Cost'] = result['Gross Cost'].map('${:,.2f}'.format)
    result['Net Cost'] = result['Net Cost'].map('${:,.2f}'.format)
    result['Total Impressions'] = result['Total Impressions'].map('{:20,.0f}'.format)

    result.replace(to_replace = r'#HHT# \d{7} - ',value = "", regex=True, inplace = True)
    result.replace(to_replace = r'#hht# \d{7} - ',value = "", regex=True, inplace = True)


    result.to_excel(r'C:\Users\%s\Downloads\Fyllo_Tremor_Usage_%s_%s.xlsx'%(user,month,year), sheet_name = 'Fyllo',index = False)
    os.remove(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\Fyllo_RX_crosstab_%s.xlsx"%(user,month))
    os.remove(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\Fyllo_crosstab_%s.xlsx"%(user,month))

def eyeota_report(name):
    user = name


    eyeota = pd.read_excel(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\Eyeota_crosstab_%s.xlsx"%(user, month))
    eyeota_rx = pd.read_excel(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\Eyeota_RX_crosstab_%s.xlsx"%(user,month))


    eyeota["Impressions"] = eyeota["Matched Impressions"] +eyeota["Modeled Impressions"]
    eyeota["Cost"] = eyeota["Total_Cost"]
    eyeota_rx["Agency"] =""
    eyeota_rx["Advertiser"] =""


    eyeota_report = eyeota.loc[:,["Agency", "Advertiser","Ext. Provider ID", "Segment ","Impressions", "Cost" ]]
    eyeota_rx_report = eyeota_rx.loc[:,["Agency", "Advertiser","External Provider ID", "Display Name","Impressions", "Cost" ]]


    eyeota_report.drop(eyeota_report.index[0],inplace = True)
    eyeota_report.drop(eyeota_report[eyeota_report["Cost"]==0].index,inplace = True)

    eyeota_rx_report.drop(eyeota_rx_report.index[0],inplace = True)
    eyeota_rx_report.drop(eyeota_rx_report[eyeota_rx_report["Cost"]==0].index,inplace = True)


    eyeota_rx_report.rename(columns = {"External Provider ID": "Segment ID","Display Name":"Segment" }, inplace = True)
    eyeota_report.rename(columns = {"Ext. Provider ID":"Segment ID","Segment ":"Segment"}, inplace =True)



    eyeota_frames = [eyeota_report,eyeota_rx_report]
    final_eyeota_report = pd.concat(eyeota_frames)

    final_eyeota_report.replace(to_replace = r'#HHT# \d{7} - ',value = "", regex=True, inplace = True)
    final_eyeota_report.replace(to_replace = r'#hht# \d{7} - ',value = "", regex=True, inplace = True)


    final_eyeota_report.to_excel(r'C:\Users\%s\Downloads\Eyeota_Tremor_Usage_%s%s.xlsx'%(user,month,year), sheet_name = 'Eyeota',index = False)
    os.remove(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\Eyeota_crosstab_%s.xlsx"%(user, month))
    os.remove(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\Eyeota_RX_crosstab_%s.xlsx"%(user,month))

def dstillery_report(name):
    user = name

    df = pd.read_excel(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\DStillery_crosstab_%s.xlsx"%(user,month))
    #Deletes first row
    report = df.drop(df.index[0])

    #Deletes any rows with 0 total cost
    report.drop(report[report['Total_Cost'] == 0].index, inplace = True)

    #Formatting the file
    report['Matched_Cost'] = report['Matched_Cost'].map('${:,.2f}'.format)
    report['Modeled_Cost'] = report['Modeled_Cost'].map('${:,.2f}'.format)
    report['Matched Impressions'] = report['Matched Impressions'].map('{:20,.0f}'.format)
    report['Modeled Impressions'] = report['Modeled Impressions'].map('{:20,.0f}'.format)
    report['Total_Cost'] = report['Total_Cost'].map('${:,.2f}'.format)
    report['Start Date'] = current_date
    report['End Date'] = date.today().replace(day=1) - timedelta(days=1)

    final_report = report.loc[:,["Buy Type","Campaign ID","Campaign Name","Placement ID",
    	"Agency","Advertiser","Ext. Provider ID","Segment ID ","Segment ","Start Date","End Date",
        "Matched Impressions","Matched_Cost","Modeled Impressions","Modeled_Cost","Total_Cost"]]

    final_report.to_excel(r'C:\Users\pambrosio\Downloads\DStillery_Tremor_Usage_%s%s.xlsx'%(month,year), sheet_name = 'DStillery',index = False)
    os.remove(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\DStillery_crosstab_%s.xlsx"%(user,month))

def adobe_report(name):
    user = name


    adobe_dsp = pd.read_excel(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\Adobe_crosstab_%s.xlsx"%(user, month))
    adobe_rx = pd.read_excel(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\Adobe_RX_Segment_Usage_crosstab_%s.xlsx"%(user, month))
    adobe_AU = pd.read_excel(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\Australia_crosstab_%s.xlsx"%(user, month))

    adobe_dsp = adobe_dsp.drop(adobe_dsp.index[0])
    adobe_rx = adobe_rx.drop(adobe_rx.index[0])
    adobe_AU = adobe_AU.drop(adobe_AU.index[0])

    adobe_AU.rename(columns =  {'Unnamed: 3': 'Segment Activated', 'Unnamed: 1': 'Campaign'}, inplace = True)

    adobe_dsp.rename(columns = {'Month of Event Date':'Month','Ext. Provider ID':'Adobe ID','Segment ':'Segment Activated','Campaign Name': 'Campaign','Total Impressions': 'Impressions'}, inplace = True)
    adobe_rx.rename(columns = {'Day of Event Date':'Month','Segment ID ':'Adobe ID','Display Name':'Segment Activated','CPM':'Data partner CPM','Cost':'Revenue payable to Data Partner'}, inplace = True)

    adobe_dsp['Month'] = f'{month} {year}'
    adobe_rx['Month'] = f'{month} {year}'

    adobe_rx['Data partner CPM'] =''
    adobe_rx['Revenue payable to Data Partner']='' 

    adobe_dsp['Data partner CPM'] =''
    adobe_dsp['Revenue payable to Data Partner']='' 

    adobe_rx["Advertiser"] = ''
    adobe_rx['Campaign'] = ''

    adobe_AU['Month']=''
    adobe_AU['Adobe ID']=''
    adobe_AU['Impressions']=''
    adobe_AU['Advertiser']=''
    adobe_AU['Data partner CPM']=''
    adobe_AU['Revenue payable to Data Partner' ]=''



    adobe_dsp_report = adobe_dsp.loc[:,['Month','Segment Activated','Adobe ID','Impressions','Advertiser','Campaign','Data partner CPM','Revenue payable to Data Partner' ]]
    adobe_rx_report = adobe_rx.loc[:,['Month','Segment Activated','Adobe ID','Impressions','Advertiser','Campaign','Data partner CPM','Revenue payable to Data Partner' ]]
    adobe_AU_report = adobe_AU.loc[:,['Month','Segment Activated','Adobe ID','Impressions','Advertiser','Campaign','Data partner CPM','Revenue payable to Data Partner' ]]


    adobe_rx_report['set index'] = ''
    adobe_AU_report['set index'] = ''


    for x in range(1,len(adobe_rx_report)+1):
        adobe_rx_report['set index'][x] = (x+len(adobe_dsp_report))

    adobe_rx_report.set_index('set index',inplace=True)

    for x in range(1,len(adobe_AU_report)+1):
        adobe_AU_report['set index'][x] = (x+len(adobe_dsp_report)+len(adobe_rx_report))

    adobe_AU_report.set_index('set index',inplace=True)



    from math import nan


    adobe_AU_report['Segment Activated'].replace(nan,'',inplace=True)

    y = len(adobe_AU_report)+len(adobe_dsp_report)+len(adobe_rx_report)
    while y > len(adobe_dsp_report)+len(adobe_rx_report):
        for x in range(len(adobe_dsp_report)+1,len(adobe_rx_report)+1):
            if adobe_rx_report['Segment Activated'][x].lower() == adobe_AU_report['Segment Activated'][y].lower():
                adobe_rx_report['Campaign'].loc[x] = adobe_AU_report['Campaign'].loc[y]
                y-=1
                continue
        else:
            y-=1



    adobe_rx_report_final = adobe_rx_report.groupby(["Month","Segment Activated","Adobe ID","Advertiser","Campaign"]).agg({"Impressions": 'sum','Data partner CPM': 'max', 'Revenue payable to Data Partner':'sum'  }, inplace = True)
    adobe_rx_report_final.reset_index(inplace = True)


    adobe_rx_report_final['set index'] = ''
    for x in range(0,len(adobe_rx_report_final)+1):
        adobe_rx_report_final['set index'][x] = (x+1+len(adobe_dsp_report))
    adobe_rx_report_final.set_index('set index',inplace=True)




    from pandas import concat

    adobe_frames = [adobe_dsp_report,adobe_rx_report_final]
    adobe_report = concat(adobe_frames)
    adobe_report



    adobe_cpm = pd.DataFrame(columns = ['Data partner CPM'])

    for x in range(1,len(adobe_report)+1):

        if adobe_report['Segment Activated'][x].find('News Corp')!=-1:
            adobe_cpm.loc[x] = 1.5
        elif adobe_report['Segment Activated'][x].find('news corp') !=-1:
            adobe_cpm.loc[x] = 1.5


        elif adobe_report['Segment Activated'][x].find('> T > flybuys')!=-1:
            adobe_cpm.loc[x] = 4
        elif adobe_report['Segment Activated'][x].find('> t > flybuys') !=-1:
            adobe_cpm.loc[x] = 4


        elif adobe_report['Segment Activated'][x].find('> D > flybuys')!=-1:
            adobe_cpm.loc[x] = 1.5
        elif adobe_report['Segment Activated'][x].find('> d > flybuys') !=-1:
            adobe_cpm.loc[x] = 1.5


        elif adobe_report['Segment Activated'][x].find('> B > flybuys')!=-1:
            adobe_cpm.loc[x] = 6
        elif adobe_report['Segment Activated'][x].find('> b > flybuys') !=-1:
            adobe_cpm.loc[x] = 6


        elif adobe_report['Segment Activated'][x].find('Experian')!=-1:
            adobe_cpm.loc[x] = 2.85
        elif adobe_report['Segment Activated'][x].find('experian')!=-1:
            adobe_cpm.loc[x] = 2.85


        elif adobe_report['Segment Activated'][x].find('Near >')!=-1:
            adobe_cpm.loc[x] = 5.5
        elif adobe_report['Segment Activated'][x].find('near >')!=-1:
            adobe_cpm.loc[x] = 5.5

        elif adobe_report['Segment Activated'][x].find('ProductReview >')!=-1:
            adobe_cpm.loc[x] = 3
        elif adobe_report['Segment Activated'][x].find('productreview >')!=-1:
            adobe_cpm.loc[x] = 3


    adobe_report['Data partner CPM'] = adobe_cpm

    import math
    cpm = adobe_report.loc[:,'Data partner CPM']
    impressions = adobe_report.loc[:,'Impressions']
    adobe_report['Revenue payable to Data Partner'] = cpm*impressions/1000


    cpm_parameter_sheet = pd.DataFrame(
    [['News Segments',1.5],
    ["Flybuys Demographic & Lifestage ",1.5],
    ['Flybuys Transactional',4],
    ['Flybuys Bespoke',6],
    ['Near',5.5],
    ['Product Review',3],
    ['Experian',2.85]],
    columns = ['Segments','CPM'])


    news_corp= adobe_report['Segment Activated'].str.contains('> News Corp >',case = False)==True
    adobe_newscorp = pd.DataFrame(columns=['Month','Segment Activated','Adobe ID','Impressions','Advertiser','Campaign','Data partner CPM','Revenue payable to Data Partner'])
    for x in range(1, len(adobe_report)+1):
        if news_corp[x] == True:
            adobe_newscorp.loc[x] = adobe_report.loc[x]


    flybuy= adobe_report['Segment Activated'].str.contains('> flybuys >',case = False)==True
    adobe_flybuy = pd.DataFrame(columns=['Month','Segment Activated','Adobe ID','Impressions','Advertiser','Campaign','Data partner CPM','Revenue payable to Data Partner'])
    for x in range(1, len(adobe_report)+1):
        if flybuy[x] == True:
            adobe_flybuy.loc[x] = adobe_report.loc[x]


    experian= adobe_report['Segment Activated'].str.contains('> Experian >',case = False)==True
    adobe_experian = pd.DataFrame(columns=['Month','Segment Activated','Adobe ID','Impressions','Advertiser','Campaign','Data partner CPM','Revenue payable to Data Partner'])
    for x in range(1, len(adobe_report)+1):
        if experian[x] == True:
            adobe_experian.loc[x] = adobe_report.loc[x]


    near= adobe_report['Segment Activated'].str.contains('> Near >',case = False)==True
    adobe_near = pd.DataFrame(columns=['Month','Segment Activated','Adobe ID','Impressions','Advertiser','Campaign','Data partner CPM','Revenue payable to Data Partner'])
    for x in range(1, len(adobe_report)+1):
        if near[x] == True:
            adobe_near.loc[x] = adobe_report.loc[x]


    product_review= adobe_report['Segment Activated'].str.contains('> ProductReview >',case = False)==True
    adobe_product_review = pd.DataFrame(columns=['Month','Segment Activated','Adobe ID','Impressions','Advertiser','Campaign','Data partner CPM','Revenue payable to Data Partner'])
    for x in range(1, len(adobe_report)+1):
        if product_review[x] == True:
            adobe_product_review.loc[x] = adobe_report.loc[x]


    import math
   

    total = (math.fsum(list(adobe_product_review['Revenue payable to Data Partner']))+
        math.fsum(list(adobe_near['Revenue payable to Data Partner']))+
        math.fsum(list(adobe_experian['Revenue payable to Data Partner']))+
        math.fsum(list(adobe_flybuy['Revenue payable to Data Partner']))+
        math.fsum(list(adobe_newscorp['Revenue payable to Data Partner'])))

    totals_sheet = pd.DataFrame(
        [['News Data',f"${math.fsum(list(adobe_newscorp['Revenue payable to Data Partner']))}"],
        ["Flybuys Demographic & Lifestage ",f"${math.fsum(list(adobe_flybuy['Revenue payable to Data Partner']))}"],
        ['Experian',f"${math.fsum(list(adobe_experian['Revenue payable to Data Partner']))}"],
        ['Near',f"${math.fsum(list(adobe_near['Revenue payable to Data Partner']))}"],
        ['Product Review',f"${math.fsum(list(adobe_product_review['Revenue payable to Data Partner']))}"],
        ['Grand Total',f"${total}"
        ]],

        
    
        columns = ['Source','Total']
    )


    with pd.ExcelWriter(r'C:/Users/%s/Downloads/NewsConnect_Usage_%s.xlsx'%(user,month)) as writer:
        totals_sheet.to_excel(writer, sheet_name='Totals')
        adobe_newscorp.to_excel(writer, sheet_name='News Connect')
        adobe_flybuy.to_excel(writer, sheet_name='flybuys')
        adobe_experian.to_excel(writer, sheet_name='Experian')
        adobe_near.to_excel(writer, sheet_name='Near Data')
        adobe_product_review.to_excel(writer, sheet_name='Product Review Data')
        cpm_parameter_sheet.to_excel(writer, sheet_name='CPM Parameters')

    os.remove(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\Adobe_crosstab_%s.xlsx"%(user, month))
    os.remove(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\Adobe_RX_Segment_Usage_crosstab_%s.xlsx"%(user, month))
    os.remove(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\Australia_crosstab_%s.xlsx"%(user, month))

def icx_report(name):

    user = name


    icx = pd.read_excel(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\ICX_crosstab_%s.xlsx"%(user,month))
    icx_rx = pd.read_excel(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\ICX_RX_crosstab_%s.xlsx"%(user,month))

    #Deletes first total row
    icx.drop(icx.index[0],inplace = True)
    icx_rx.drop(icx_rx.index[0],inplace = True)

    #Deletes rows with $0 total cost
    icx.drop(icx[icx['Total_Cost'] == 0].index, inplace = True)
    icx_rx.drop(icx_rx[icx_rx['Cost'] == 0].index, inplace = True)

    #Adding a "Provider Column"
    icx["Provider"] = "ICX"
    icx_rx["Provider"] = "ICX"

    #Deleting HHT labels
    icx.replace(to_replace = r'#HHT# \d{7} - ',value = "", regex=True, inplace = True)
    icx_rx.replace(to_replace = r'#HHT# \d{7} - ',value = "", regex=True, inplace = True)


    #New report does not have start and end dates for RX. Setting them to the 1st and last day of the month
    icx["Start Date"] = "{}/{}/{}".format(prev.strftime("%m"),prev.replace(day=1).day,prev.strftime("%Y"))
    icx["End Date"] = "{}/{}/{}".format(prev.strftime("%m"),prev.day,prev.strftime("%Y"))

    icx_rx["Start Date"] = "{}/{}/{}".format(prev.strftime("%m"),prev.replace(day=1).day,prev.strftime("%Y"))
    icx_rx["End Date"] = "{}/{}/{}".format(prev.strftime("%m"),prev.day,prev.strftime("%Y"))


    #Reordering the columns
    icx_rx_final = icx_rx.loc[:,["Start Date", 
                                "End Date",
                                "Provider", 
                                "Display Name",
                                "External Provider ID", 
                                "Impressions",
                                "CPM",
                                "Cost"]]

    icx_final = icx.loc[:,["Start Date", 
                        "End Date",
                        "Provider",  
                        "Segment ",
                        "Ext. Provider ID", 
                        "Total Impressions",
                        "Segment CPM",
                        "Total_Cost"]]

    #Renaming the columns
    icx_rx_final.rename(columns = {"Display Name":"Segment", 
                                "Display Name": "Segment",
                                "External Provider ID": "IDs",
                                "Cost": "Price"}, 
                                inplace = True)
    icx_final.rename(columns = {"Segment ": "Segment", 
                                "Ext. Provider ID": "IDs", 
                                "Total Impressions": "Impressions",
                                "Segment CPM": 'CPM',
                                "Total_Cost": "Price"}, 
                                inplace = True)

    #Combining RX and DSP
    icx_frames = [icx_final, icx_rx_final]
    final_icx_report = pd.concat(icx_frames)

    final_icx_report.to_excel(r'C:\Users\%s\Downloads\ICX_Tremor_Usage_%s_%s.xlsx'%(user,month,year), sheet_name = 'ICX',index = False)
    os.remove(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\ICX_crosstab_%s.xlsx"%(user,month))
    os.remove(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\ICX_RX_crosstab_%s.xlsx"%(user,month))

def neustar_report(name):
    user = name

    neustar = pd.read_excel(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\Neustar_crosstab_%s.xlsx"%(user,month))

    neustar.drop(neustar.index[0],inplace = True)

    neustar["Cost"] = neustar["Matched Impressions"]*neustar["Segment CPM"]/1000

    neustar_report = neustar.loc[:,["Agency","Advertiser","Ext. Provider ID","Segment ","Matched Impressions","Segment CPM","Cost"]]
    neustar_report.rename(columns = {"Segment ":"Segment Name","Matched Impressions":"Impressions", "Segment CPM":"CPM"},inplace = True)
    neustar_report.drop(neustar_report.index[0],inplace = True)
    neustar_report.drop(neustar_report[neustar_report['Cost']==0].index, inplace = True)


    #Formatting so output has commas, $, etc.
    neustar_report['Cost'] = neustar_report['Cost'].map('${:,.2f}'.format)
    neustar_report['CPM'] = neustar_report['CPM'].map('${:,.2f}'.format)
    neustar_report['Impressions'] = neustar_report['Impressions'].map('{:20,.0f}'.format)

    neustar_report.to_excel(r'C:\Users\%s\Downloads\Neustar_Tremor_Usage_%s%s.xlsx'%(user,month,year), sheet_name = 'Neustar',index = False)
    os.remove(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\Neustar_crosstab_%s.xlsx"%(user,month))

def comscoretv_report(name):
    user = name


    comscore = pd.read_excel(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\ComscoreTV_crosstab_%s.xlsx"%(user,month))
    comscore_rx = pd.read_excel(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\ComscoreTV_RX_crosstab_%s.xlsx"%(user,month))

    #Comscore DSP
    comscore['Platform (TV DSP or UnrulyX)'] = "TV DSP"
    comscore["Event_date"]=str(current_date)

    #Comscore RX
    comscore_rx["Event_date"]=str(current_date)
    comscore_rx['Platform (TV DSP or UnrulyX)'] = "UnrulyX"

    #Formatting
    comscore_report = comscore.loc[:,["Event_date","Platform (TV DSP or UnrulyX)","Ext. Provider ID","Segment ","Agency","Advertiser","Total Impressions", "Segment CPM","Total_Cost"]]

    #Formatting: Deleting 1st total row
    comscore_report.drop(comscore_report.index[0],inplace = True)
    comscore_rx.drop(comscore_rx.index[0],inplace = True)


    #Deleting rows with $0 total cost
    comscore_report.drop(comscore_report[comscore_report['Total_Cost'] == 0].index, inplace = True)
    comscore_rx.drop(comscore_rx[comscore_rx['Cost'] == 0].index, inplace = True)

    #Pivoting to combine all the dates into 1 row
    comscore_final = comscore_report.groupby(["Event_date","Platform (TV DSP or UnrulyX)","Ext. Provider ID","Segment ","Agency","Advertiser"]).agg({"Total Impressions": 'sum',"Segment CPM": 'max', "Total_Cost":'sum'  }, inplace = True)
    comscore_rx_report = comscore_rx.groupby(["Event_date","Platform (TV DSP or UnrulyX)","External Provider ID","Segment"]).agg({"Impressions": 'sum',"CPM": 'max', "Cost":'sum'  }, inplace = True)

    #Re-indexing to ensure all item labels are repeated
    comscore_rx_report.reset_index(inplace = True)
    comscore_final.reset_index(inplace = True)

    #Formatting so output has commas, $, etc.
    comscore_final['Total_Cost'] = comscore_final['Total_Cost'].map('${:,.2f}'.format)
    comscore_final['Segment CPM'] = comscore_final['Segment CPM'].map('${:,.2f}'.format)
    comscore_final['Total Impressions'] = comscore_final['Total Impressions'].map('{:20,.0f}'.format)

    #Formatting so output has commas, $, etc.
    comscore_rx_report['Cost'] = comscore_rx_report['Cost'].map('${:,.2f}'.format)
    comscore_rx_report['CPM'] = comscore_rx_report['CPM'].map('${:,.2f}'.format)
    comscore_rx_report['Impressions'] = comscore_rx_report['Impressions'].map('{:20,.0f}'.format)

    #Inserting blank agency and advertiser columns into the RX report
    comscore_rx_report["Agency"] = ""
    comscore_rx_report["Advertiser"] = ""
    comscore_rx_final = comscore_rx_report.loc[:,["Event_date","Platform (TV DSP or UnrulyX)","External Provider ID","Segment","Agency","Advertiser","Impressions", "CPM","Cost"]]

    #Renaming columns to match naming convention
    comscore_final.rename(columns={"Event_date":"Date","Ext. Provider ID":"Segment ID","Segment ":"Attribute/Segment","Total Impressions":"Impressions","Segment CPM":"CPM","Total_Cost":"Total Cost"}, inplace = True)
    comscore_rx_final.rename(columns={"Event_date":"Date","External Provider ID":"Segment ID","Segment":"Attribute/Segment", "Cost":"Total Cost"}, inplace = True)

    #Combining DSP and RX reports
    comscore_frames = [comscore_final, comscore_rx_final]
    final_comscore_report = pd.concat(comscore_frames)

    #Removing HHT labels
    final_comscore_report.replace(to_replace = r'#HHT# \d{7} - ',value = "", regex=True, inplace = True)
    final_comscore_report.replace(to_replace = r'#hht# \d{7} - ',value = "", regex=True, inplace = True)

    #Sending report to downloads
    final_comscore_report.to_excel(r'C:\Users\pambrosio\Downloads\Comscore TV_Tremor_Usage_%s%s.xlsx'%(month,year), sheet_name = 'Comscore TV',index = False)

    #removing files
    os.remove(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\ComscoreTV_crosstab_%s.xlsx"%(user,month))
    os.remove(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\ComscoreTV_RX_crosstab_%s.xlsx"%(user,month))

def comscorepa_report(name):
    user = name

    context_comscore = pd.read_excel(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\ComscorePA_crosstab_%s.xlsx"%(user,month))
    context_comscore_rx = pd.read_excel(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\ComscorePA_RX_crosstab_%s.xlsx"%(user,month))

    #Comscore DSP
    context_comscore['Platform (TV DSP or UnrulyX)'] = "TV DSP"
    context_comscore["Date"]=str(current_date)

    #Comscore RX
    context_comscore_rx["Date"]=str(current_date)
    context_comscore_rx['Platform (TV DSP or UnrulyX)'] = "UnrulyX"
    context_comscore_rx["Agency"] = ""
    context_comscore_rx["Advertiser"] = ""
    context_comscore_rx["Net after 20% rev share to Unruly/Tremor"] = context_comscore_rx["Cost"]*.8

    #Formatting
    context_comscore_report = context_comscore.loc[:,["Date","Platform (TV DSP or UnrulyX)","External_Provider_Id","Segment Name","Agency","Advertiser","Matched Impressions", "CPM (copy)","Matched Cost","RevShare Cost (copy)"]]
    context_comscore_rx_report = context_comscore_rx.loc[:,["Date","Platform (TV DSP or UnrulyX)","Segment ID","Display Name","Agency","Advertiser","Impressions", "CPM","Cost","Net after 20% rev share to Unruly/Tremor"]]
    context_comscore_report.rename(columns={"External_Provider_Id":"Segment ID","Segment Name":"Attribute/Segment", "Matched Impressions":"Impressions","CPM (copy)":"CPM","Matched Cost":"Total Cost","RevShare Cost (copy)": "Net after 20% rev share to Unruly/Tremor"}, inplace = True)
    context_comscore_rx_report.rename(columns={"Display Name":"Attribute/Segment", "Cost":"Total Cost"}, inplace = True)

    #Deleting rows with $0 total cost
    context_comscore_report.drop(context_comscore_report[context_comscore_report['Total Cost'] == 0].index, inplace = True)
    context_comscore_rx_report.drop(context_comscore_rx_report[context_comscore_rx_report['Total Cost'] == 0].index, inplace = True)

    #Deleting HHT labels
    context_comscore_report.replace(to_replace = r'#HHT# \d{7} - ',value = "", regex=True, inplace = True)
    context_comscore_rx_report.replace(to_replace = r'#HHT# \d{7} - ',value = "", regex=True, inplace = True)

    #Pivoting to combine all the dates into 1 row
    context_comscore_final = context_comscore_report.groupby(["Date","Platform (TV DSP or UnrulyX)","Segment ID","Attribute/Segment","Agency","Advertiser"]).agg({"Impressions": 'sum',"CPM": 'max', "Total Cost":'sum', "Net after 20% rev share to Unruly/Tremor":'sum'  }, inplace = True)
    context_comscore_rx_final = context_comscore_rx_report.groupby(["Date","Platform (TV DSP or UnrulyX)","Segment ID","Attribute/Segment","Agency","Advertiser"]).agg({"Impressions": 'sum',"CPM": 'max', "Total Cost":'sum',"Net after 20% rev share to Unruly/Tremor":'sum'  }, inplace = True)

    #Re-indexing to ensure all item labels are repeated
    context_comscore_rx_final.reset_index(inplace = True)
    context_comscore_final.reset_index(inplace = True)

    #Formatting so output has commas, $, etc.
    context_comscore_final['Total Cost'] = context_comscore_final['Total Cost'].map('${:,.2f}'.format)
    context_comscore_final["Net after 20% rev share to Unruly/Tremor"] = context_comscore_final["Net after 20% rev share to Unruly/Tremor"].map('${:,.2f}'.format)
    context_comscore_final['CPM'] = context_comscore_final['CPM'].map('${:,.2f}'.format)
    context_comscore_final['Impressions'] = context_comscore_final['Impressions'].map('{:20,.0f}'.format)

    #Formatting so output has commas, $, etc.
    context_comscore_rx_final['Total Cost'] = context_comscore_rx_final['Total Cost'].map('${:,.2f}'.format)
    context_comscore_rx_final['CPM'] = context_comscore_rx_final['CPM'].map('${:,.2f}'.format)
    context_comscore_rx_final['Impressions'] = context_comscore_rx_final['Impressions'].map('{:20,.0f}'.format)

    #Inserting blank agency and advertiser columns into the RX report
    context_comscore_rx_final_report = context_comscore_rx_final.loc[:,["Date","Platform (TV DSP or UnrulyX)","Segment ID","Attribute/Segment","Agency","Advertiser","Impressions", "CPM","Total Cost","Net after 20% rev share to Unruly/Tremor"]]

    #Combining DSP and RX reports
    context_comscore_frames = [context_comscore_final, context_comscore_rx_final]
    context_final_comscore_report = pd.concat(context_comscore_frames)

    context_final_comscore_report.to_excel(r'C:\Users\pambrosio\Downloads\Comscore Contextual_Tremor_Usage_%s_%s.xlsx'%(month,year), sheet_name = 'Comscore Contextual',index = False)

    #removing files
    os.remove(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\ComscorePA_crosstab_%s.xlsx"%(user,month))
    os.remove(r"C:\Users\%s\OneDrive - Tremor International\Documents\Automated_Usage_Reports\attachments\ComscorePA_RX_crosstab_%s.xlsx"%(user,month))