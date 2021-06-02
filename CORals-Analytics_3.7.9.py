#!/usr/bin/env python3.7

# vim: set fileencoding="utf-8":

#CORals Analytics v.3.7.9
#This script is created by Christian Otto Ruge and CORals.
#It is licenced under GNU GPL v.3.
#Github: https://github.com/christianoruge/corals_analytics/blob/main/CORals-Analytics_3.7.9.py
#https://www.corals.no



import os
import pandas as pd
import csv
import csv23
import xlsxwriter
from xlsxwriter import Workbook
import pingouin as pg
import xlrd
import PySimpleGUI as sg
from io import open
import seaborn as sns
import matplotlib.pyplot as plt
from scipy import stats
from scipy.stats import pearsonr
import statsmodels.api as sm
from statsmodels.stats.outliers_influence import variance_inflation_factor
import statsmodels.formula.api as smf
from pyprocessmacro import Process
from factor_analyzer import FactorAnalyzer
from factor_analyzer import ConfirmatoryFactorAnalyzer
from factor_analyzer import ModelSpecificationParser


sg.theme('Light Grey 1')


#Popup for exceptions:
def popup_break(event, message):
    while True:
        sg.Popup(message)
        event=='False'
        break

layoutOriginal = [
    [sg.Text('')],
    [sg.Text('CORALS ANALYTICAL TOOLS', size=(25,1), justification='left', font=("Arial", 20))],
    [sg.Text('')],
    [sg.Text('')],  
    [sg.Text('Choose a CORals-tool', key='Choose', font=('bold'))],     
    [sg.Frame(layout=[      
    [sg.Radio('CSV - rescue', "RADIO1", key="CSV", default=False, size=(20,1)), sg.Radio('Distribution', "RADIO1", key="Distribution", default=False, size=(20,1)), sg.Radio('Correlation', "RADIO1", key="Correlation",  default=False, size=(20,1))],         
    [sg.Radio('Regression', "RADIO1", key="Regression", default=False, size=(20,1)), sg.Radio('Mediation', "RADIO1", key="Mediation", default=False, size=(20,1)), sg.Radio('Moderation', "RADIO1", key="Moderation", default=False, size=(20,1))],
    [sg.Radio('Factor analysis', "RADIO1", key="Factor", default=False, size=(20,1)), sg.Radio('Create scales', "RADIO1", key="Scales", default=False, size=(20,1))]], title='Options',title_color='red', relief=sg.RELIEF_SUNKEN, tooltip='Use these to set flags')],      
    [sg.Text('')],
    [sg.Text('')],  
    [sg.Button('Continue'), sg.Button('Close')],
    [sg.Text('')],
    [sg.Image('logo.png', key='icon', size=(450, 200))],
    [sg.Text('')],
    [sg.Text('©Licenced under GNU GPL v2 by Christian Otto Ruge', justification='right')]]

try:    
    winOriginal = sg.Window('CORals Analytics v. 3.7.9', default_element_size=(40, 1), grab_anywhere=False, location=(100,100), size=(530,650)).Layout(layoutOriginal)    

    winCsv_active=False
    winCorrelation_active=False
    winRegression_active=False
    winMediation_active=False
    winModeration_active=False
    winDistribution_active=False
    winFactor_active=False
    winScales_active=False
    

    while True:
        evOriginal,valOriginal=winOriginal.Read(timeout=100)
        if evOriginal is None or evOriginal=='Close':
            winOriginal.Close()
            del winOriginal
            break

        if (not winCsv_active) and (valOriginal['CSV']==True) and (evOriginal=='Continue'):
            winOriginal.Hide()
            winCsv_active=True


            layoutCsv = [
                [sg.Text('')], 
                [sg.Text('CSV CONVERTER:', size=(25,1), justification='left', font=("Arial", 20))],
                [sg.Text('')],
                [sg.Text('')], 
                [sg.Text('Choose CSV-file:', font=('bold'))],
                [sg.In('', key='csv-file', size=(60,1)), sg.FileBrowse()],
                [sg.Radio('Encoding 1 (PC)', 'RADIO1', key='win', default=True, size=(20,1)),  sg.Radio('Encoding 2 (Mac)', 'RADIO1', key='mac', default=False, size=(20,1))],  
                [sg.Text('')],
                [sg.Text('NB: Csv-files can be coded in various ways, the above')],
                [sg.Text('alternatives represent two... Hope one does the job for you!')],
                [sg.Text('Select output folder:', size=(35, 1), font=('bold'))],      
                [sg.Text('')],
    
                [sg.InputText('', key='xlsx_output', size=(60,1)), sg.FolderBrowse()],
                [sg.Button('Fix and convert'), sg.Button('Back')],
                [sg.Text('')],
                [sg.Text('')],
                
                
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('©Licenced under GNU GPL v2 by Christian Otto Ruge', justification='right')]]
            

            winCsv=sg.Window('CORals Analytics v. 3.7.9', default_element_size=(40, 1), grab_anywhere=False, location=(100,100), size=(530,650)).Layout(layoutCsv)

            while True:
                evCsv, valCsv = winCsv.Read(timeout=100)
                
                if evCsv is None or evCsv == 'Back':
                    winCsv_active=False
                    winCsv.Close()
                    del winCsv
                    winOriginal.UnHide()
                    break
                
                if (evCsv=='Fix and convert') and (valCsv['csv-file']==''):
                    popup_break(evCsv, 'Choose csv-file')

                if (evCsv=='Fix and convert') and (valCsv['xlsx_output']==''):
                    popup_break(evCsv, 'Choose output folder')

                if not (valCsv['csv-file']=='') and not (valCsv['xlsx_output']=='') and (evCsv=='Fix and convert'):   
                    while True:
                        dataset = valCsv['csv-file']
                        filename=os.path.basename(dataset)
                        csv_file=str(filename)
                        
                        
                        folder=dataset.replace(csv_file,'')

                        output_folder=valCsv['xlsx_output']
                    

                        #file=[]
                                

                        if valCsv['mac']==True:
                            

                            try:
                                file=[]
                                
                                pd_dataset=pd.read_csv(dataset)
                                    
                                delimiter_check=str(pd_dataset)
                                
                                
                                if ";" in delimiter_check:
                                    delimiter=';'
                                else:
                                    delimiter=','
                                

                                
                                with open(dataset, 'r', encoding='iso-8859-1') as g:
                                    reader=csv.reader(g, delimiter=delimiter)
                                    for row in reader:
                                        file.append(row)

                            except:
                                popup_break(evCsv, 'There is an error. Please retry with "Generated on MAC" checked')
                                break
                            

                            file_clean=[] #Norwegian letters from win
                            #missing for lower case 'æ' and 'å'.
                            
                            for row in file:
                                
                                #row=[w.replace('≈','æ') for w in row]
                                row=[w.replace('Â¯','ø') for w in row]
                                #row=[w.replace('≈','å') for w in row]
                                row=[w.replace('âˆ†','Æ') for w in row]
                                row=[w.replace('Ã¿','Ø') for w in row]
                                row=[w.replace('â‰ˆ','Å') for w in row]
                                
                                row=[w.replace('¯','ø') for w in row]
                                row=[w.replace('Ê','æ') for w in row]
                                row=[w.replace('Â','å') for w in row]
                                row=[w.replace('ÿ','Ø') for w in row]
                                row=[w.replace('∆','Æ') for w in row]
                                
                                row=[w.replace('‚âà','Å') for w in row]
                                row=[w.replace('√ø','Ø') for w in row]
                                row=[w.replace('¬Ø','ø') for w in row]
                                row=[w.replace('‚àÜ','Æ') for w in row]
                                row=[w.replace('â','Å') for w in row]
                                
                                
                                file_clean.append(row)

                            
                               
                            
                            df=pd.DataFrame(file_clean)

                            

                            xlsx_filename=csv_file.replace(".csv", ".xlsx")

                            engine = 'xlsxwriter'

                            with pd.ExcelWriter(os.path.join(output_folder, xlsx_filename), engine=engine) as writer:
                                df.to_excel(writer, sheet_name="Fixed file", index = None, header=False)

                            break 
                                
                            
                            
                            
                                

                        if valCsv['win']==True:
                            print('PC')
                            
                            try:
                            
                                file=[]
                            
                                with open(dataset, 'r', encoding='iso-8859-1') as g:
                                    for row in g:
                                        file.append(row)
                                print(file)
                                file=[c.replace(';', ',') for c in file]
                                file=[c.replace('"', '') for c in file]

                            except:
                                popup_break(evCsv, 'There is an error. Please retry with "Generated on PC" checked')
                                break
                    
                            file_split=[]

                            for row in file:
                                split_row=row.split(',')
                                file_split.append(split_row)
                            
                            df=pd.DataFrame(file_split)

                            print(df)

                            xlsx_filename=csv_file.replace(".csv", ".xlsx")

                            engine = 'xlsxwriter'

                            with pd.ExcelWriter(os.path.join(output_folder, xlsx_filename), engine=engine) as writer:
                                df.to_excel(writer, sheet_name="Fixed file", index = None, header=False)

                            break
                            
                        

                        


                        
                        file_split=[]
                        for row in file:
                            split_row=row.split(',')
                            file_split.append(split_row)

                        Print('Filesplit ok')
                        
                        df=pd.DataFrame(file_split)

                        print(df)

                        xlsx_filename=csv_file.replace(".csv", ".xlsx")

                        engine = 'xlsxwriter'

                        with pd.ExcelWriter(os.path.join(output_folder, xlsx_filename), engine=engine) as writer:
                            df.to_excel(writer, sheet_name="Fixed file", index = None, header=False)

                        break


    #Correlation
        
        if (not winCorrelation_active) and (valOriginal['Correlation']==True) and (evOriginal=='Continue'):
            winOriginal.Hide()
            winCorrelation_active=True


            layoutCorrelation = [
                [sg.Text('')],
                [sg.Text('BIVARIATE CORRELATION:', size=(25,1), justification='left', font=("Arial", 20))],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('Choose dataset (xlsx-file):', font=('bold'))],
                [sg.In('', key='correlation_dataset', size=(60,1)), sg.FileBrowse()],
                [sg.Text('Enter variables to be analyzed:', font=('bold'))],
                [sg.Text('(Variable names are the column headers (row 1) in the xlsx-file.)')],
                [sg.Text('Divide variables by one comma only (no space).')],
                [sg.InputText('', key='correlation_variables', size=(60,1))],
                [sg.Text('')],
                [sg.Text('Select output formats:', font=('bold'))],
                [sg.Checkbox('Correlation matrix', key='matrix', default=True, size=(20,1)),  sg.Checkbox('Correlation heatmap', key='heatmap', default=True, size=(20,1))],  
                [sg.Text('Select output folder:', font=('bold'), size=(60,1))],      
                [sg.InputText('', key='correlation_output', size=(60,1)), sg.FolderBrowse()],
                [sg.Button('Continue'), sg.Button('Back')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('©Licenced under GNU GPL v2 by Christian Otto Ruge', justification='right')]]

            
            winCorrelation=sg.Window('CORals Analytics v. 3.7.9', default_element_size=(40, 1), grab_anywhere=False, location=(100,100), size=(530,650)).Layout(layoutCorrelation)

            while True:
                evCorrelation, valCorrelation = winCorrelation.Read(timeout=100)    

                if evCorrelation in (None, 'Back'):
                    winCorrelation_active=False
                    winCorrelation.Close()
                    del winCorrelation
                    winOriginal.UnHide()
                    break       


                if (evCorrelation=='Continue' and valCorrelation['correlation_dataset']==''):
                    popup_break(evCorrelation, 'Choose dataset')
                    
                if (evCorrelation=='Continue' and valCorrelation['correlation_variables']==''):
                    popup_break(evCorrelation, 'Choose variables')
                    
                #Valdidate variables supplied
                if (evCorrelation=='Continue') and not (valCorrelation['correlation_dataset']=='') and not (valCorrelation['correlation_variables']==''):
                    dataset = valCorrelation['correlation_dataset']
                    data=pd.read_excel(dataset)
                    validvariables=list(data.columns)
                    var=valCorrelation['correlation_variables']
                    varsplit=var.split(',')
                    list_var=list(varsplit)
                    
                    response='yes'
                    for string in list_var:    
                        if string not in validvariables:
                            response='no'

                            
                    if (evCorrelation=='Continue') and (response =='no'):
                        popup_break(evCorrelation, 'Variable(s) not in dataset')

                    #End validation of variables
                    
                    if (evCorrelation=='Continue' and valCorrelation['correlation_output']==''):
                        popup_break(evCorrelation, 'Choose output folder')   




                    if (evCorrelation=='Continue') and not (valCorrelation['correlation_dataset']=='') and not (valCorrelation['correlation_variables']=='') and not (valCorrelation['correlation_output']=='') and (response=='yes'):
                        dataset = valCorrelation['correlation_dataset']
                        data=pd.read_excel(dataset)
                        validvariables=list(data.columns)
                        var=valCorrelation['correlation_variables']
                        varsplit=var.split(',')
                        list_var=list(varsplit)
                        
                        while True:
                            chosendata=data[varsplit]
                            output_folder=valCorrelation['correlation_output']

                            if valCorrelation['heatmap']==True:

                                #Create and save heatmap as png-file
                                
                                corrs = chosendata.corr()
                                mask = np.zeros_like(corrs)
                                mask[np.triu_indices_from(mask)] = True
                                fig=plt.figure()
                                heatmap=sns.heatmap(corrs, cmap='Spectral_r', mask=mask, square=True, vmin=-.5, vmax=.5)

                                varnames=str(varsplit)
                                defname=varnames.replace('[','')
                                defname=varnames.replace(']','')

                                plt.title('CORRELATION-MATRIX (HEATMAP):')
                                b, t = plt.ylim() # discover the values for bottom and top (Thanks to SalMac86, GitHub, 25.10.19)
                                b += 0.5 # Adds 0.5 to the bottom
                                t -= 0.5 # Subtracts 0.5 from the top
                                plt.ylim(b, t) # updates the ylim(bottom, top) values
                                plt.tight_layout()
                                fig.savefig(os.path.join(output_folder, 'correlation-heatmap_') + var+'.png', dpi=400)
                                fig=None


                            if valCorrelation['matrix']==True:

                                #Create and save correlation matrix as xlsx-file
                                plot=pd.DataFrame(chosendata.rcorr(stars=True))
                                description=pd.DataFrame([['CORRELATION MATRIX:'],
                                [''],
                                ['Upper triangle: p-values'],
                                ['Lower triangle: Pearson r-values:'],
                                ['']])
                                
                                engine = 'xlsxwriter'
                                with pd.ExcelWriter(os.path.join(output_folder, 'correlation-plot_') +var+'.xlsx', engine=engine) as writer:
                                    description.to_excel(writer, sheet_name="pearson r and p-values", index = None, header=True, startrow=0)
                                    plot.to_excel(writer, sheet_name="pearson r and p-values", index = None, header=True, startrow= 6)

                            break


                   
    #Regression
        if (not winRegression_active) and (valOriginal['Regression']==True) and (evOriginal=='Continue'):
            winOriginal.Hide()
            winRegression_active=True

            layoutRegression = [
                [sg.Text('')],
                [sg.Text('REGRESSION ANALYSIS:', justification='left', font=('Arial', 20))],
                [sg.Text('')],
                [sg.Text('Choose dataset (xlsx-file):', font=('bold'))],
                [sg.In('', key='regression_dataset', size=(60,1)), sg.FileBrowse()],
                [sg.Text('')],
                [sg.Text('Choose regression type,', font=('bold'))],
                [sg.Text('simple (for 1 IV) or multiple (for multiple IVs).')],
                [sg.Radio('Simple', 'RADIO1', key='Simple', default=False, size=(20,1)), sg.Radio('Multiple', 'RADIO1', key='Multiple', default=False, size=(20,1))],
                [sg.Text('Enter independent variable(s)', font=('bold'))],
                [sg.Text('(Column headers (row 1) in the xlsx-file function as variable names.)')],
                [sg.Text('For multiple regressin, separate IVs by commas only (no space).')],
                [sg.InputText('', key='iv', size=(60,1))],
                [sg.Text('Enter dependent variable (DV).', font=('bold'))],
                [sg.InputText('', key='dv', size=(15,1))],
                [sg.Text('')],
                [sg.Text('Select output formats:', font=('bold'))],
                [sg.Checkbox('Regression plots', key='regression_plots', default=True, size=(16,1)), sg.Checkbox('Table', key='table', default=True, size=(16,1)), sg.Checkbox('VIF values (for multiple regression only)', key='vif', default=False, size=(40,1))],
                [sg.Text('Select output folder:', font=('bold'))],
                [sg.InputText('', key='regression_output', size=(60,1)), sg.FolderBrowse()],
                [sg.Button('Continue'), sg.Button('Back')],
                [sg.Text('')],
                [sg.Text('©Licenced under GNU GPL v2 by Christian Otto Ruge', justification='right')]]

            winRegression=sg.Window('CORals Analytics v. 3.7.9', default_element_size=(40, 1), grab_anywhere=False, location=(100,100), size=(530,650)).Layout(layoutRegression)
  
            while True:
                evRegression, valRegression = winRegression.Read(timeout=100)
                
                if evRegression in (None, 'Back'):
                    winRegression_active=False
                    winRegression.Close()
                    del winRegression
                    winOriginal.UnHide()
                    break

                if (evRegression=='Continue' and valRegression['regression_dataset']==''):
                    popup_break(evRegression, 'Please select dataset')
                    
                if (evRegression=='Continue') and valRegression['Simple']==False and valRegression['Multiple']==False:
                    popup_break(evRegression, 'Regresion type unselected.')

                if (evRegression=='Continue' and valRegression['iv']==''):
                    popup_break(evRegression, 'Please select independent variable(s)')

                if (evRegression=='Continue') and (',' in valRegression['iv']) and (valRegression['Simple']==True):
                    popup_break(evRegression, 'NB: Only one independent variable in single regression analysis.')

                if (evRegression=='Continue' and valRegression['dv']==''):
                    popup_break(evRegression, 'Please select dependent variable')
                    
                if (evRegression=='Continue') and (',' in valRegression['dv']):
                    popup_break(evRegression, 'Please select only one dependent variable.')

                if (evRegression=='Continue') and (valRegression['regression_plots']==False) and (valRegression['table']==False) and (valRegression['vif']==False):
                    popup_break(evRegression, 'Please select output format')


                if (evRegression=='Continue' and valRegression['regression_output']==''):
                    popup_break(evRegression, 'Please select output folder')

                
                #Valdidate variables supplied
                if (evRegression=='Continue') and not (valRegression['regression_dataset']=='') and not (valRegression['iv']=='') and not (valRegression['dv']=='') and not (',' in valRegression['dv']):
                    dataset = valRegression['regression_dataset']
                    data=pd.read_excel(dataset)
                    validvariables=list(data.columns)
                    var=valRegression['iv']+','+valRegression['dv']
                    varsplit=var.split(',')
                    list_var=list(varsplit)
                    
                    response='yes'
                    for string in list_var:    
                        if string not in validvariables:
                            response='no'
 
                    
                    if (evRegression=='Continue') and (response =='no'):
                        popup_break(evRegression, 'Cancel. One or multiple variable(s) not in dataset')

                    #End validation of variables
                        
                    if (evRegression=='Continue') and not (valRegression['regression_output']=='') and (response=='yes'):


                        iv=valRegression['iv']
                        iv_str="'"+iv+"'"
                        ivsplit=iv.split(',')#list of multiple IVs for multiple regression
                        dv=valRegression['dv']
                        simple=valRegression['Simple']
                        multiple=valRegression['Multiple']
                        output_folder=valRegression['regression_output']
                        X=data[ivsplit]              
                        y=data[dv]
                        X=sm.add_constant(X)
                        model=sm.OLS(y, X).fit()
                        
                        predictions=model.predict(X)



                        if simple==True:

                            if valRegression['vif']==True:
                                popup_break(evRegression, 'Please uncheck "VIF - values" for single regression')
            
                            if valRegression['table']==True:#Make and save table:

                                fig=plt.figure(figsize=(12, 7))

                                plt.text(0.01, 0.05, str(model.summary()), {'fontsize': 10}, fontproperties = 'monospace')
                                plt.axis('off')
                                plt.savefig(os.path.join(output_folder, 'Simple_OLS_')+iv+'_'+dv+'_table.png', dpi=300, format='png', transparent=True)
                                plt.savefig(os.path.join(output_folder, 'Simple_OLS_')+iv+'_'+dv+'_table.pdf', dpi=300, format='pdf')

                            if valRegression['regression_plots']==True:#Make and save regression plots:

                                fig, ax = plt.subplots(figsize=(19,10))
                                plt.axis('off')
                                fig=sm.graphics.plot_regress_exog(model, iv, fig=fig)
                                
                                plt.savefig(os.path.join(output_folder, 'Simple_OLS_')+iv+'_'+dv+'_regression_plots.png', dpi=300, format='png', transparent=True)
                                plt.savefig(os.path.join(output_folder, 'Simple_OLS_')+iv+'_'+dv+'_regression_plots.pdf', dpi=300, format='pdf')

                        elif multiple==True:

                            if valRegression['table']==True: #Save table (model summary):
                                
                                print(model.summary())##Used with large amounts of variables.
                                fig=plt.figure(figsize=(12,7))
                                plt.text(0.01, 0.05, str(model.summary()), {'fontsize': 10}, fontproperties = 'monospace')
                                plt.axis('off')
                                fig.savefig(os.path.join(output_folder, 'Multiple_OLS_')+iv + '_'+dv+'_table.png', dpi=300, format='png', transparent=True)
                                fig.savefig(os.path.join(output_folder, 'Multiple_OLS_')+iv + '_'+dv+'_table.pdf', dpi=300, format='pdf', transparent=True)
                                
                                
                            if valRegression['regression_plots']==True: #Make and save ccpr plots:
                                fig, ax = plt.subplots(figsize=(19,10))
                                plt.axis('off')
                                fig=sm.graphics.plot_ccpr_grid(model, fig=fig)
                                
                                plt.savefig(os.path.join(output_folder, 'Multiple_OLS_')+iv+'_'+dv+'_ccpr_plots.png', dpi=300, transparent=True)

                            if valRegression['vif']==True: #Make and save VIF values:
                                    
                                vif = pd.DataFrame()
                                vif["VIF Factor"] = [variance_inflation_factor(X.values, i) for i in range(X.shape[1])]
                                vif["Variables"] = X.columns
                                form=vif.round(1)

                                df=pd.DataFrame(form)
                                engine = 'xlsxwriter'

                                with pd.ExcelWriter(os.path.join(output_folder, 'Multiple_OLS_')+iv+'_'+dv+'_VIF_values.xlsx', engine=engine) as writer:
                                    df.to_excel(writer, sheet_name="VIF values", index = None, header=True)
                                    

    #Mediation
        if not winMediation_active and valOriginal['Mediation']==True and evOriginal=='Continue':
            winOriginal.Hide()
            winMediation_active=True

            layoutMediation = [
                [sg.Text('')],
                [sg.Text("MEDIATION ANALYSIS", size=(25,1), justification='left', font=("Arial", 20))],
                [sg.Text('Based on the Process Macro by Andrew F. Hayes, Ph.D. (www.afhayes.com)')],
                [sg.Text("The following analysis is equv. to Hayes' Process, model 4:")],
                [sg.Text('')],
                [sg.Text('Choose dataset (xlsx-file):', font=('bold'))],
                [sg.InputText('', key='mediation_dataset', size=(60,1)), sg.FileBrowse(target='mediation_dataset')],
                [sg.Text('Use the column headers (row 1) in the xlsx-file as variable names.')],
                [sg.Text('')],
                [sg.Text('Type independent variable name (single):', font=('bold'))],
                [sg.InputText('', key='iv', size=(15,1))],
                [sg.Text('')],
                [sg.Text('Type mediating variable name (M):', font=('bold'))],
                [sg.InputText('', key='m', size=(15,1))],
                [sg.Text('')],
                [sg.Text('Type dependent variable name (DV):', font=('bold'))],
                [sg.InputText('', key='dv', size=(15,1))],
                [sg.Text('')],
                [sg.Text('Select output folder:', size=(35, 1), font=('bold'))],      
                [sg.InputText('', key='mediation_output', size=(60,1)), sg.FolderBrowse(target='mediation_output')],           
                [sg.Button('Continue'), sg.Button('Back')],
                [sg.Text('')],
                [sg.Text('©Licenced under GNU GPL v2 by Christian Otto Ruge', justification='right')]]
            


            winMediation=sg.Window('CORals Analytics v. 3.7.9', default_element_size=(40, 1), grab_anywhere=False, location=(100,100), size=(530,650)).Layout(layoutMediation)


            while True:
                evMediation, valMediation = winMediation.Read(timeout=100)
                response='none'
                

                #Exceptions:

                if evMediation is None or evMediation == 'Back':
                    winMediation_active=False
                    winMediation.Close()
                    del winMediation
                    winOriginal.UnHide()
                    break
                
                if (evMediation=='Continue' and valMediation['mediation_dataset']==''):
                    popup_break(evMediation, 'Please select dataset')
                    
                if (evMediation=='Continue' and valMediation['iv']==''):
                    popup_break(evMediation, 'Please select independent variable')

                if (evMediation=='Continue') and (',' in valMediation['iv']):
                    popup_break(evMediation, 'Please select only one independent variable')
                        
                if (evMediation=='Continue' and valMediation['m']==''):
                    popup_break(evMediation, 'Please select mediating variable')

                if (evMediation=='Continue') and (',' in valMediation['m']):
                    popup_break(evMediation, 'Please select only one mediating variable')

                if (evMediation=='Continue' and valMediation['dv']==''):
                    popup_break(evMediation, 'Please select dependent variable')
                        
                if (evMediation=='Continue') and (',' in valMediation['dv']):
                    popup_break(evMediation, 'Please select only one mediating variable')

                if (evMediation=='Continue' and valMediation['mediation_output']==''):
                    popup_break(evMediation, 'Please select output folder')
                
                #Valdidate variables supplied
                if (evMediation=='Continue') and not (valMediation['mediation_dataset']=='') and not (valMediation['iv']=='') and not (valMediation['m']=='') and not (valMediation['dv']==''):
                    
                    dataset = valMediation['mediation_dataset']
                    data=pd.read_excel(dataset)
                    validvariables=list(data.columns)
                    var=valMediation['iv']+','+valMediation['m']+','+valMediation['dv']
                    varsplit=var.split(',')
                    list_var=list(varsplit)
                    
                    response='yes'
                    for string in list_var:    
                        if string not in validvariables:
                            response='no'

                            
                    if (evMediation=='Continue') and (response =='no'):
                        response='none'
                        popup_break(evMediation, 'One or several entered variables not in dataset')
                    
                    if (evMediation=='Continue') and (response=='yes') and not (valMediation['mediation_output']==''):

                        iv=valMediation['iv']
                        dv=valMediation['dv']
                        med=valMediation['m']
            
                        data=pd.read_excel(dataset)
                        output_folder=valMediation['mediation_output']
                            
                        med_dataframe=pd.DataFrame()

                        fig = plt.figure(figsize=(8,8))
                        p = Process(data=data, model=4, x=iv, y=dv, m=med)
                        model_mediation=p.outcome_models[med]
                        csum_med=model_mediation.coeff_summary()
                        msum_med=model_mediation.model_summary()

                        med_dataframe=med_dataframe.append(csum_med)

                        direct_model=p.direct_model
                        csum_dir=direct_model.coeff_summary()

                        header_const_uv=pd.DataFrame(columns=["Coefficients (constant and iv):"])
                        header_direct=pd.DataFrame(columns=["Direct effect:"])
                        header_ind=pd.DataFrame(columns=["Indirect effect:"])
                        header_modelsummary=pd.DataFrame(columns=["Model summary:"])
                       

                        indirect_model=p.indirect_model
                        csum_ind=indirect_model.coeff_summary()
                        csum_ind_list=csum_ind.values.tolist()
            
                        
                        #Plotting
                        p.plot_conditional_direct_effects(med)

                        engine = 'xlsxwriter'

                        

                        with pd.ExcelWriter(os.path.join(output_folder, 'Mediation_')+iv+'_'+med+'_'+dv+'.xlsx', engine=engine) as writer:
                            header_modelsummary.to_excel(writer, sheet_name="Model_summary", index = None, startrow=0, startcol=2, header=True)
                            msum_med.to_excel(writer, sheet_name="Model_summary", index = None, startrow=1, header=True)
                            header_const_uv.to_excel(writer, sheet_name="Model_summary", index = None, startrow=5, startcol=2, header=True)
                            csum_med.to_excel(writer, sheet_name="Model_summary", index = None, startrow=6, header=True)
                            header_direct.to_excel(writer, sheet_name="Model_summary", index = None, startrow=11, startcol=2, header=True)
                            csum_dir.to_excel(writer, sheet_name="Model_summary", index = None, startrow=12, header=True)
                            header_ind.to_excel(writer, sheet_name="Model_summary", index = None, startrow=16, startcol=2, header=True)
                            csum_ind.to_excel(writer, sheet_name="Model_summary", index = None, startrow=17, header=True)

                
    #Moderation
        if not winModeration_active and valOriginal['Moderation']==True and evOriginal=='Continue':
            winOriginal.Hide()
            winModerarion_active=True

            layoutModeration = [
                [sg.Text('')],
                [sg.Text("MODERATION ANALYSIS", size=(25,1), justification='left', font=("Arial", 20))],
                [sg.Text('Based on the Process Macro by Andrew F. Hayes, Ph.D. (www.afhayes.com)')],
                [sg.Text("The following analysis is equv. to Hayes' Process, model 1:")],
                [sg.Text('')],
                [sg.Text('Choose dataset (xlsx-file):', font=('bold'))],
                [sg.InputText('', key='moderation_dataset', size=(60,1)), sg.FileBrowse(target='moderation_dataset')],
                [sg.Text('Use the column headers (row 1) in the xlsx-file as variable names.')],
                [sg.Text('Type independent variable name (single):', font=('bold'))],

                [sg.InputText('', key='iv', size=(15,1))],
                [sg.Text('Type moderating variable name (M) (NB: dictonomous):', font=('bold'))],
                [sg.InputText('', key='m', size=(15,1))],
                [sg.Text('Type dependent variable name (DV):', font=('bold'))],
                [sg.InputText('', key='dv', size=(15,1))],
                [sg.Text('')],
                [sg.Text('Choose output format:', font=('bold'))],
                [sg.Checkbox('Export values to Excel', key='xlsx', default=True, size=(22,1)),  sg.Checkbox('Save line chart as pdf', key='pdf', default=True, size=(22,1)), sg.Checkbox('Save line chart as png', key='png', default=True, size=(30,1))],  
                [sg.Text('')],
                [sg.Text('Select output folder:', size=(35, 1), font=('bold'))],      
                [sg.InputText('', key='moderation_output', size=(60,1)), sg.FolderBrowse(target='moderation_output')],           
                [sg.Button('Continue'), sg.Button('Back')],
                [sg.Text('')],
                [sg.Text('©Licenced under GNU GPL v2 by Christian Otto Ruge', justification='right')]]
            

            winModeration=sg.Window('CORals Analytics v. 3.7.9', default_element_size=(40, 1), grab_anywhere=False, location=(100,100), size=(530,650)).Layout(layoutModeration)

            while True:
                evModeration, valModeration = winModeration.Read(timeout=100)
                response='none'
 
                #Exceptions:

                if evModeration is None or evModeration == 'Back':
                    winModeration_active=False
                    winModeration.Close()
                    del winModeration
                    winOriginal.UnHide()
                    break

                
                if (evModeration=='Continue' and valModeration['moderation_dataset']==''):
                    popup_break(evModeration, 'Please select dataset')

                      
                if (evModeration=='Continue' and valModeration['iv']==''):
                    popup_break(evModeration, 'Please select independent variable')
                    
                if (evModeration=='Continue' and ',' in valModeration['iv']):
                    popup_break(evModeration, 'Please select only one independent variable')
                        
                if (evModeration=='Continue' and valModeration['m']==''):
                    popup_break(evModeration, 'Please select moderating variable')
                    
                if (evModeration=='Continue' and ',' in valModeration['m']):
                    popup_break(evModeration, 'Please select only one moderating variable')
                    
                if (evModeration=='Continue' and valModeration['dv']==''):
                    popup_break(evModeration, 'Please select dependent variable')
                    
                if (evModeration=='Continue' and ',' in valModeration['dv']):
                    popup_break(evModeration, 'Please select only one dependent variable')
                
                if (evModeration=='Continue' and valModeration['moderation_output']==''):
                    popup_break(evModeration, 'Please select output folder')


                 #Valdidate variables supplied
                if (evModeration=='Continue') and not (valModeration['iv']=='') and not (',' in valModeration['iv']) and not (valModeration['m']=='') and not (',' in valModeration['m']) and not (valModeration['dv']=='') and not (',' in valModeration['dv']) and not (valModeration['moderation_dataset']=='') and not (valModeration['moderation_output']==''):
                    print('4')
                    dataset = valModeration['moderation_dataset']
                    data=pd.read_excel(dataset)
                    iv=valModeration['iv']
                    dv=valModeration['dv']
                    med=valModeration['m']
                    output_folder=valModeration['moderation_output']
                    validvariables=list(data.columns)
                    var=valModeration['iv']+','+valModeration['m']+','+valModeration['dv']
                    varsplit=var.split(',')
                    list_var=list(varsplit)
                    
                    #response='yes'
                    for string in list_var:    
                        if string not in validvariables:
                            response='no'
                            
                        else:
                            response='yes'

                    print(response)

                    if response=='no':
                        response='none'
                        popup_break(evModeration, 'One or several entered variables not in dataset')

                    moderatorvalues=[]
                    for i in data[med]:
                        if i not in moderatorvalues:
                            moderatorvalues.append(i)


                    if not len(moderatorvalues)==2:
                        popup_break(evModeration, 'Moderating variable is not dictonomous (only two values)')
                                        
                    if (evModeration=='Continue') and (response=='yes') and len(moderatorvalues)==2:


                        p = Process(data=data, model=1, x=iv, y=dv, m=med)
                        direct_model=p.direct_model
                        direct_summary=direct_model.coeff_summary()

                        if valModeration['xlsx']==True:
                            
                            header=pd.DataFrame(columns=["Results from the moderation analysis (the direct effect):"])

                            engine = 'xlsxwriter'

                            with pd.ExcelWriter(os.path.join(output_folder, 'Moderation_')+iv+'_'+med+'_'+dv+'_moderation.xlsx', engine=engine) as writer:
                                header.to_excel(writer, sheet_name="Results", startcol=1, startrow=0)
                                direct_summary.to_excel(writer, sheet_name="Results", startcol=-1, startrow=1, header=True)


                            data=pd.read_excel(os.path.join(output_folder, 'Moderation_')+iv+'_'+med+'_'+dv+'_moderation.xlsx')
                            fig = plt.figure(figsize=(8,8))
                            plt.title('Moderation analysis: '+med)

                            


                        x=np.linspace(1,5,num=5)
                        a1=data.iloc[1,-6]
                        a2=data.iloc[2,-6]

                        y1=a1*x
                        y2=a2*x

                        ax = fig.add_subplot(111)
                        ax.set_xlabel('Level of '+iv)
                        ax.set_ylabel('Level of '+dv)
                        ax.spines['right'].set_color('none')
                        ax.spines['top'].set_color('none')

                        plt.plot(x, y1, 'b', label='Category: '+str(data.iloc[1,-7]), linewidth=3)
                        plt.plot(x, y2, 'r', label='Category: '+str(data.iloc[2,-7]), linewidth=3)
                        plt.axhline(y=0, color='k', linestyle='-', linewidth=1)
                        plt.axhline(y=1,color='k', linestyle=':', linewidth=0.5)
                        plt.axhline(y=2,color='k', linestyle=':', linewidth=0.5)
                        plt.axhline(y=3,color='k', linestyle=':', linewidth=0.5)
                        plt.axhline(y=-1,color='k', linestyle=':', linewidth=0.5)
                        plt.axhline(y=-2,color='k', linestyle=':', linewidth=0.5)
                        plt.axhline(y=-3,color='k', linestyle=':', linewidth=0.5)

                        plt.legend(loc='upper left')
                        plt.ylim((-3, 3))
                        plt.xlim((1, 5))
                        plt.tight_layout()
                        x=np.linspace(1,5,5)
                        ax.spines['right'].set_color('none')
                        ax.spines['top'].set_color('none')

                        

                        if valModeration['pdf']==True:
                            plt.savefig(os.path.join(output_folder, 'Moderation_')+iv+'_'+med+'_'+dv+'_moderation.pdf', dpi=300, format='pdf')
                        if valModeration['png']==True:
                            plt.savefig(os.path.join(output_folder, 'Moderation_')+iv+'_'+med+'_'+dv+'_moderation.png', dpi=300, format='png', transparent=True)

    #Descriptive
        
        if (not winDistribution_active) and (valOriginal['Distribution']==True) and (evOriginal=='Continue'):
            winOriginal.Hide()
            winDistribution=True


            layoutDistribution = [
                [sg.Text('')],
                [sg.Text('DISTRIBUTION ANALYSES:', size=(25,1), justification='left', font=("Arial", 20))],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('Choose dataset (xlsx-file):', font=('bold'))],
                [sg.In('', key='distribution_dataset', size=(60,1)), sg.FileBrowse()],
                [sg.Text('Enter variables to be analyzed:', font=('bold'))],
                [sg.Text('(Variable names are the column headers (row 1) in the xlsx-file.)')],
                [sg.Text('Divide variables by one comma only (no space).')],
                [sg.InputText('', key='distribution_variables', size=(60,1))],
                [sg.Text('Optional: Enter hue-variables (for pairwise plots only)', font=('bold'))],
                [sg.Text('(Hue variable must be one of the above chosen variables.))')],
                [sg.InputText('', key='hue', size=(60,1))],
                [sg.Text('')],
                [sg.Text('Select output formats:', font=('bold'))],
                [sg.Checkbox('Line charts', key='lines', default=True, size=(20,1)),  sg.Checkbox('Bars', key='bars', default=True, size=(20,1))],  
                [sg.Checkbox('Pairwise plots', key='plots', default=True, size=(20,1)), sg.Checkbox('Descriptive values', key='values', default=True, size=(20,1))], 
                [sg.Text('Select output folder:', font=('bold'), size=(35, 1))],      
                [sg.InputText('', key='distribution_output', size=(60,1)), sg.FolderBrowse()],
                [sg.Button('Continue'), sg.Button('Back')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('©Licenced under GNU GPL v2 by Christian Otto Ruge', justification='right')]]

            
            winDistribution=sg.Window('CORals Analytics v. 3.7.9', default_element_size=(40, 1), grab_anywhere=False, location=(100,100), size=(530,650)).Layout(layoutDistribution)
        
            
            while True:
                evDistribution, valDistribution = winDistribution.Read(timeout=100)
        
                
                if evDistribution in (None, 'Back'):
                    winDistribution_active=False
                    winDistribution.Close()
                    del winDistribution
                    winOriginal.UnHide()
                    break       


                if (evDistribution=='Continue') and (valDistribution['distribution_dataset']==''):
                    popup_break(evDistribution, 'Choose dataset')
                    
                if (evDistribution=='Continue' and valDistribution['distribution_variables']==''):
                    popup_break(evDistribution, 'Choose variables')

                if (evDistribution=='Continue') and (not valDistribution['hue']=='') and ((valDistribution['lines']==True) or (valDistribution['values']==True) or (valDistribution['bars']==True)) and (valDistribution['plots']==False):
                    popup_break(evDistribution, 'Enter hue value only for "pairwise plots"')

                if (evDistribution=='Continue') and (valDistribution['lines']==False) and (valDistribution['bars']==False) and (valDistribution['plots']==False) and (valDistribution['values']==False):
                    popup_break(evDistribution, 'Choose output format')

                if (evDistribution=='Continue') and (',' in valDistribution['hue']):
                    popup_break(evDistribution, 'select only one hue variable')
                                                                             

                if (evDistribution=='Continue') and (not valDistribution['hue']==''):
                    color=valDistribution['hue']
                    

                if (evDistribution=='Continue') and (valDistribution['hue']==''):
                    color=None
                    
                #Valdidate variables supplied
                if (evDistribution=='Continue') and not (valDistribution['distribution_dataset']=='') and not (valDistribution['distribution_variables']==''):
                    dataset = valDistribution['distribution_dataset']
                    data=pd.read_excel(dataset)
                    validvariables=list(data.columns)
                    var=valDistribution['distribution_variables']
                    varnames=str(var)
                    varnames=varnames.replace('[','')
                    varnames=varnames.replace(']','')
                    varsplit=var.split(',')
                    list_var=list(varsplit)
                    response='yes'
                    print(color)

                    if (evDistribution=='Continue') and (not valDistribution['hue']=='') and (not valDistribution['hue'] in list_var):   
                        popup_break(evDistribution, 'Hue variable must be one of the above entered variables.')
                    
                    for string in list_var:    
                        if string not in validvariables:
                            response='no'
                    
                    if (evDistribution=='Continue') and (response =='no'):
                        popup_break(evDistribution, 'Variable(s) not in dataset')              
                    
                    if (evDistribution=='Continue') and (valDistribution['distribution_output']==''):
                        popup_break(evDistribution, 'Choose output folder')   


                    if (evDistribution=='Continue') and (not valDistribution['distribution_output']=='') and (response=='yes'):                                                
                        print('true')    
                        while True:
                            
                            output_folder=valDistribution['distribution_output']
                            df=data[varsplit]

                            if valDistribution['lines']==True:

                                number=0
                                for n in list_var:
                                    for row in df[n]:
                                        if row>number:
                                            number=row

                                    count  = df[n].value_counts()
                                    count = count[:,]
                                    plt.figure(figsize=(10,5))
                                    sns.lineplot(count.index, count.values, alpha=0.8, linewidth=2.5, marker="o")
                                    plt.title('Distribution of '+n)
                                    plt.ylabel('Number of observations', fontsize=12)
                                    plt.xlabel(n, fontsize=12)
                                    plt.savefig(os.path.join(output_folder, 'Line_')+n+'_distribution.png', dpi=300, format='png', transparent=True)

                            if valDistribution['bars']==True:
                                number=0
                                for n in list_var:
                                    for row in df[n]:
                                        if row>number:
                                            number=row

                                    count  = df[n].value_counts()
                                    count = count[:,]
                                    plt.figure(figsize=(10,5))
                                    sns.barplot(count.index, count.values, alpha=0.8)
                                    plt.title('Distribution of '+n)
                                    plt.ylabel('Number of observations', fontsize=12)
                                    plt.xlabel(n, fontsize=12)
                                    plt.savefig(os.path.join(output_folder, 'Bars_')+n+'_distribution.png', dpi=300, format='png', transparent=True)
                               

                            if valDistribution['plots']==True:
                                sns.pairplot(df, diag_kind="auto", hue=color, palette="husl")
                                
                                plt.savefig(os.path.join(output_folder, 'Pair_plots_')+varnames+'_distribution.png', dpi=300, format='png', transparent=True) 

                            if valDistribution['values']==True:
                                values=df.describe()                             
                                values.to_excel(os.path.join(output_folder, 'Values_')+varnames+'.xlsx')


                            #break
    #Factor analyses
    #https://factor-analyzer.readthedocs.io/en/latest/factor_analyzer.html
                        
        if not winFactor_active and valOriginal['Factor']==True and evOriginal=='Continue':
            winOriginal.Hide()
            winFactor_active=True

            layoutFactor = [
                
                [sg.Text("FACTOR ANALYSIS", size=(25,1), justification='left', font=("Arial", 20))],

                [sg.Text('Choose dataset (xlsx-file):', font=('bold'))],
                [sg.InputText('', key='factor_dataset', size=(60,1)), sg.FileBrowse(target='factor_dataset')],
                [sg.Text('Select output folder:', size=(35, 1), font=('bold'))],
                [sg.InputText('', key='factor_output', size=(60,1)), sg.FolderBrowse(target='factor_output')],         
                [sg.Text('Select desired analyses and outputs:', font=('bold'))],
                [sg.Text('Exploratory factor analysis (EFA):', font=('bold'))],
                [sg.Text('Type variable names (divide by","):', font=('bold'))],
                [sg.Text('Use the column headers (row 1) in the xlsx-file as variable names.')],

                [sg.InputText('', key='variables', size=(60,1))],        
                [sg.Text('Select number of factors:     '), sg.Slider(key='n_factors', range=(2,6 ), orientation='h', size=(20,20), default_value=3)],              
                [sg.Checkbox('Eigenvalues', key='eigenvalues', default=False, size=(20,1)), sg.Checkbox('Factor analysis without rotation', key='factor_unrotated', default=False, size=(30,1))],  

                [sg.Checkbox('Varimax rotation', key='varimax', default=False, size=(20,1)), sg.Checkbox('Promax rotation', key='promax', default=False, size=(20,1))],
                [sg.Text('Confirmatory factor analysis (CFA):', font=('bold'))],
                [sg.Text('Select up to factor names and belonging variable names (divided by",").')],
                [sg.Text('Factor name:', size=(8,1)), sg.InputText('', key='factor_name_1', size=(15,1)), sg.Text('Variables:', size=(7,1)), sg.InputText('', key='factor_variables_1', size=(30,1))],
                [sg.Text('Factor name:', size=(8,1)), sg.InputText('', key='factor_name_2', size=(15,1)), sg.Text('Variables:', size=(7,1)), sg.InputText('', key='factor_variables_2', size=(30,1))],
                [sg.Text('Factor name:', size=(8,1)), sg.InputText('', key='factor_name_3', size=(15,1)), sg.Text('Variables:', size=(7,1)), sg.InputText('', key='factor_variables_3', size=(30,1))],
                [sg.Text('Factor name:', size=(8,1)), sg.InputText('', key='factor_name_4', size=(15,1)), sg.Text('Variables:', size=(7,1)), sg.InputText('', key='factor_variables_4', size=(30,1))],
                [sg.Button('Continue'), sg.Button('Back')],
                [sg.Text('')],
                [sg.Text('©Licenced under GNU GPL v2 by Christian Otto Ruge', justification='right')]]

            winFactor=sg.Window('CORals Analytics v. 3.7.9', default_element_size=(40, 1), grab_anywhere=False, location=(100,100), size=(530,650)).Layout(layoutFactor)

            while True:
                evFactor, valFactor = winFactor.Read(timeout=100)
                response='none'
                
                #Exceptions:

                if evFactor is None or evFactor == 'Back':
                    winFactor_active=False
                    winFactor.Close()
                    del winFactor
                    winOriginal.UnHide()
                    break
                
                if (evFactor=='Continue' and valFactor['factor_dataset']==''):
                    popup_break(evFactor, 'Please select dataset')         

                if (evFactor=='Continue' and valFactor['variables']=='' and valFactor['factor_variables_1']==''):
                    popup_break(evFactor, 'Please select variables')

                if (evFactor=='Continue') and (' ' in valFactor['variables']):
                    popup_break(evFactor, 'Please divide variables by only "," (no space)')
                                          
                if (evFactor=='Continue' and valFactor['factor_output']==''):
                    popup_break(evFactor, 'Please select output folder')

                
                #Valdidate variables supplied
                if (evFactor=='Continue') and not (valFactor['factor_dataset']=='') and not (' ' in valFactor['variables']==''):

                                    
                    dataset = valFactor['factor_dataset']
                    data=pd.read_excel(dataset)
                    validvariables=list(data.columns)
                    list_var=''

                    if not valFactor['variables']=='':
                        
                        var=valFactor['variables']
                        varsplit=var.split(',')
                        list_var=list(varsplit)
                        variables_df=pd.DataFrame(list_var)
                        variables_df.columns=['Variables']
                        df_features=data[list_var]

                    factors_n=int(valFactor['n_factors'])
                    

                    if factors_n == 2:
                        headings=['Factor 1', 'Factor 2']
                    if factors_n == 3:
                        headings=['Factor 1', 'Factor 2', 'Factor 3']
                    if factors_n == 4:
                        headings=['Factor 1', 'Factor 2', 'Factor 3', 'Factor 4']
                    if factors_n == 5:
                        headings=['Factor 1', 'Factor 2', 'Factor 3', 'Factor 4', 'Factor 5']
                    if factors_n == 6:
                        headings=['Factor 1', 'Factor 2', 'Factor 3', 'Factor 4', 'Factor 5', 'Factor 6']

                 
                    response='yes'
                    if not list_var=='':
                        for string in list_var:    
                            if string not in validvariables:
                                response='no'                  
                    
                    if (evFactor=='Continue') and (response =='no'):
                        response='none'
                        popup_break(evFactor, 'One or several entered variables not in dataset')
                                           
                    if (evFactor=='Continue') and (response=='yes') and not (valFactor['factor_output']==''):            
                        
                        output_folder=valFactor['factor_output']
                                              
                        #EFA Factor analyses:
                        
                        if evFactor=='Continue' and valFactor['factor_unrotated']==True:
                                                       
                            fa=FactorAnalyzer(n_factors=factors_n, rotation=None)
                            fa.fit(df_features)
                            unrotated_fit=str(fa.fit(df_features))
                            unrotated_loadings=fa.loadings_
                            unrotated_communalities=fa.get_communalities()
                           
                            unrotated_loadings_df=pd.DataFrame(data=unrotated_loadings)
                            unrotated_loadings_df.columns=headings
                            unrotated_loadings_df_descr=variables_df.join(unrotated_loadings_df)
                            unrotated_communalities_df=pd.DataFrame(data=unrotated_communalities)
                            
                            unrotated_fit_list=list(unrotated_fit.split('-'))
                            unrotated_fit_df=pd.DataFrame(unrotated_fit_list)
                              
                            engine = 'xlsxwriter'
                           
                            with pd.ExcelWriter(os.path.join(output_folder, 'EFA_factor_unrotated_'+var+'.xlsx') , engine=engine) as writer:
                                unrotated_loadings_df_descr.to_excel(writer, sheet_name="Unrotated factor loadings", index = None, header=True)
                                unrotated_communalities_df.to_excel(writer, sheet_name="Unrotated factor communalities", index = None, header=True)
                                unrotated_fit_df.to_excel(writer, sheet_name="Unrotated factor fit", index = None, header=True)


                        if evFactor=='Continue' and valFactor['varimax']==True:
                                                     
                            fa=FactorAnalyzer(n_factors=factors_n, rotation='Varimax')
                            fa.fit(df_features)
                            varimax_fit=str(fa.fit(df_features))
                            varimax_loadings=fa.loadings_
                            varimax_communalities=fa.get_communalities()

                            varimax_loadings_df=pd.DataFrame(data=varimax_loadings)
                            varimax_loadings_df.columns=headings
                            varimax_loadings_df_descr=variables_df.join(varimax_loadings_df)
                            varimax_communalities_df=pd.DataFrame(data=varimax_communalities)
                            
                            varimax_fit_list=list(varimax_fit.split('-'))
                            varimax_fit_df=pd.DataFrame(varimax_fit_list)
                                                      
                            engine = 'xlsxwriter'
                                                        
                            with pd.ExcelWriter(os.path.join(output_folder, 'EFA_factor_varimax_'+var+'.xlsx') , engine=engine) as writer:
                                varimax_loadings_df_descr.to_excel(writer, sheet_name="Varimax factor loadings", index = None, header=True)
                                varimax_communalities_df.to_excel(writer, sheet_name="Varimax factor communalities", index = None, header=True)
                                varimax_fit_df.to_excel(writer, sheet_name="Varimax factor fit", index = None, header=True)
                      
                        if evFactor=='Continue' and valFactor['promax']==True:
                                                     
                            fa=FactorAnalyzer(n_factors=factors_n, rotation='Promax')
                            fa.fit(df_features)
                            promax_fit=str(fa.fit(df_features))
                            promax_loadings=fa.loadings_
                            promax_communalities=fa.get_communalities()
           
                            promax_loadings_df=pd.DataFrame(data=promax_loadings)
                            promax_loadings_df.columns=headings
                            promax_loadings_df_descr=variables_df.join(promax_loadings_df)
                            promax_communalities_df=pd.DataFrame(data=promax_communalities)
                            
                            promax_fit_list=list(promax_fit.split('-'))
                            promax_fit_df=pd.DataFrame(promax_fit_list)
                                                      
                            engine = 'xlsxwriter'
                                                        
                            with pd.ExcelWriter(os.path.join(output_folder, 'EFA_factor_promax_'+var+'.xlsx') , engine=engine) as writer:
                                promax_loadings_df_descr.to_excel(writer, sheet_name="Promax factor loadings", index = None, header=True)
                                promax_communalities_df.to_excel(writer, sheet_name="Promax factor communalities", index = None, header=True)
                                promax_fit_df.to_excel(writer, sheet_name="Promax factor fit", index = None, header=True)

                        if evFactor=='Continue' and valFactor['eigenvalues']==True:
                                                     
                            fa=FactorAnalyzer(rotation=None)
                            
                            eigenvalues_fit=fa.fit(df_features)
                            eigenvalues=fa.get_eigenvalues()
                            eigenvalues_df=pd.DataFrame(data=eigenvalues)
  
                            eigenvalues_df.columns=list_var
                            eigenvalues_df.drop(eigenvalues_df.tail(1).index,inplace=True)  # drop last row
                          
                            engine = 'xlsxwriter'
                          
                            with pd.ExcelWriter(os.path.join(output_folder, 'factor_eigenvalues_'+var+'.xlsx') , engine=engine) as writer:
                                eigenvalues_df.to_excel(writer, sheet_name="Factor eigenvalues", index = None, header=True)
                                
                        
                         #CFA Factor analyses:
                        
                       
                        if evFactor=='Continue' and not (valFactor['factor_name_1']=='') and not (valFactor['factor_variables_1']==''):
                            
                            var_cfa=str(valFactor['factor_variables_1'])+','+str(valFactor['factor_variables_2'])+','+str(valFactor['factor_variables_3'])+','+str(valFactor['factor_variables_4'])
                                
                            varsplit_cfa=var_cfa.split(',')
                            list_var_cfa=list(varsplit_cfa)
                                                
                            while '' in list_var_cfa:
                                list_var_cfa.remove('')

                            response_cfa='yes'
                            for string in list_var_cfa:    
                                if string not in validvariables:
                                    response_cfa='no'

                            if (evFactor=='Continue') and (response_cfa =='no'):
                                response_cfa='none'
                                popup_break(evFactor, 'One or several entered cfa-variables not in dataset')
                        
                    
                            if (evFactor=='Continue') and (response_cfa=='yes'):            
                 
                                factor_name_1=valFactor['factor_name_1']
                                factor_name_2=valFactor['factor_name_2']
                                factor_name_3=valFactor['factor_name_3']
                                factor_name_4=valFactor['factor_name_4']
                                factor_variables_1=valFactor['factor_variables_1'].split(',')
                                factor_variables_2=valFactor['factor_variables_2'].split(',')
                                factor_variables_3=valFactor['factor_variables_3'].split(',')
                                factor_variables_4=valFactor['factor_variables_4'].split(',')
                               
                                dataframe_features=data[list_var_cfa]

                              
                                model_dict_string_1="'"+str(factor_name_1)+"': "+str(list(factor_variables_1))
                                model_dict_string_2="'"+str(factor_name_2)+"': "+str(list(factor_variables_2))
                                model_dict_string_3="'"+str(factor_name_3)+"': "+str(list(factor_variables_3))
                                model_dict_string_4="'"+str(factor_name_4)+"': "+str(list(factor_variables_4))
                                
                                model_dict_string_tot=model_dict_string_1

                                variable_names_list=[str(factor_name_1)]
                                if not factor_name_2=='':
                                    model_dict_string_tot=model_dict_string_tot+', '+model_dict_string_2
                                    variable_names_list.append(str(factor_name_2))
                                if not factor_name_3=='':
                                    model_dict_string_tot=model_dict_string_tot+', '+model_dict_string_3
                                    variable_names_list.append(str(factor_name_3))
                                if not factor_name_4=='':
                                    model_dict_string_tot=model_dict_string_tot+', '+model_dict_string_4
                                    variable_names_list.append(str(factor_name_4))
                                
                                model_dict=eval('{'+model_dict_string_tot+'}')

                                
                                
                                model_spec = ModelSpecificationParser.parse_model_specification_from_dict(dataframe_features, model_dict)

                                cfa = ConfirmatoryFactorAnalyzer(model_spec, disp=False)
                                cfa.fit(dataframe_features.values)

                                cfa_loadings=cfa.loadings_

                                
                                cfa_factor_varcovs=cfa.factor_varcovs_
                                
                                #cfa_transforms=cfa.transform(dataframe_features.values)
                                #cfa_std_err=cfa.get_standard_errors()
                                

                                cfa_loadings_df=pd.DataFrame(data=cfa_loadings)
                                cfa_loadings_df.columns=variable_names_list

                              
                                cfa_factor_varcovs_df=pd.DataFrame(data=cfa_factor_varcovs)
                                cfa_factor_varcovs_df.columns=variable_names_list
                                
                                #NB: If you want cfa transform and std errors, comment out the following lines and the last lines in the excel-seksjen below.
                                #cfa_transforms_df=pd.DataFrame(data=cfa_transforms)
                                #cfa_transforms_df.columns=variable_names_list
                                #cfa_std_err_df=pd.DataFrame(data=cfa_std_err)
                                
                                engine = 'xlsxwriter'

                                with pd.ExcelWriter(os.path.join(output_folder, 'CFA_'+var_cfa+'.xlsx') , engine=engine) as writer:
                                    cfa_loadings_df.to_excel(writer, sheet_name="CFA loadings", index = None, header=True)
                                    cfa_factor_varcovs_df.to_excel(writer, sheet_name="CFA factor varcovs", index = None, header=True)
                                    #cfa_transforms_df.to_excel(writer, sheet_name="CFA Transforms", index = None, header=True)
                                    #cfa_std_err_df.to_excel(writer, sheet_name="CFA Std. errors", index = None, header=True)

                                
   #Scales
        if (not winScales_active) and (valOriginal['Scales']==True) and (evOriginal=='Continue'):
            winOriginal.Hide()
            
            winScales_active=True

            layoutScales = [
                [sg.Text('')],
                [sg.Text('CREATE SCALES(combined variables):', justification='left', font=('Arial', 20))],
                [sg.Text('')],
                [sg.Text('Choose dataset (xlsx-file):', font=('bold'))],
                [sg.In('', key='original_dataset', size=(60,1)), sg.FileBrowse()],
                [sg.Text('')],
                [sg.Text('Choose variables to combine:', font=('bold'))],
                
                [sg.Text('Column headers (row 1) in the xlsx-file function as variable names.')],
                [sg.Text('Separate variable names by commas only (no space):')],
                [sg.InputText('', key='variables', size=(60,1))],
                [sg.Text('Enter name of new scale/new variable:', font=('bold'))],
                [sg.InputText('', key='new_variable_name', size=(15,1))],

                [sg.Text('')],
                [sg.Text('')],
                [sg.Button('Continue'), sg.Button('Back')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('©Licenced under GNU GPL v2 by Christian Otto Ruge', justification='right')]]
            
            winScales=sg.Window('CORals Analytics v. 3.7.9', default_element_size=(40, 1), grab_anywhere=False, location=(100,100), size=(530,650)).Layout(layoutScales)
           
            while True:
                evScales, valScales = winScales.Read(timeout=100)
                

                if evScales in (None, 'Back'):
                    winScales_active=False
                    winScales.Close()
                    del winScales
                    winOriginal.UnHide()
                    break

                if (evScales=='Continue' and valScales['original_dataset']==''):
                    popup_break(evScales, 'Please select dataset')
                    
                if (evScales=='Continue') and valScales['variables']=='':
                    popup_break(evScales, 'Please select variables to combine.')

                if (evScales=='Continue' and valScales['new_variable_name']==''):
                    popup_break(evScales, 'Please select name of new scale/variable.')
                
                #Valdidate variables supplied
                if (evScales=='Continue') and not (valScales['original_dataset']=='') and not (valScales['variables']=='') and not (valScales['new_variable_name']==''):
                    dataset = valScales['original_dataset']
                    
                    data=pd.read_excel(dataset)
                    validvariables=list(data.columns)
                    
                    var=valScales['variables']
                    
                    varsplit=var.split(',')
                    list_var=list(varsplit)                   
                
                    filename=str(os.path.basename(dataset))
                    folder=dataset.replace('/'+filename,'')

                    response='yes'
                    for string in list_var:    
                        if string not in validvariables:
                            response='no'                  
                    
                    if (evScales=='Continue') and (response =='no'):
                        popup_break(evScales, 'Cancel. One or multiple variable(s) not in supplied dataset')

                    #End validation of variables
                       
                    if (evScales=='Continue') and (response=='yes'):


                        variables=valScales['variables']
                        variables="['"+variables+"']"
                        variables_column=variables.replace(",","','")
                        variables_column=variables_column.replace(' ', '')
                        variables_list=[]
                        variables_alone=valScales['variables'].replace(',',' ')
                        for word in variables_alone.split():
                            variables_list.append(word)

                        new_variable="'"+valScales['new_variable_name']+"'"                        
                        data[new_variable]=data[variables_list].mean(axis=1, skipna=True)

                        if filename.startswith('scales_incl_'):
                            new_filename=filename
                        else:
                            new_filename='scales_incl_'+filename
                
                        engine = 'xlsxwriter'

                        with pd.ExcelWriter(os.path.join(folder, new_filename) , engine=engine) as writer:
                            data.to_excel(writer, sheet_name="Dataset_modified", index = None, header=True)

                       
except:
    sg.Popup('Ooops! Something went wrong! This may be due to missing or non-numerical variables in the analyzed areas of the dataset, invalid input or another unexpected issue. If variables are missing it is recommended to delete these rows. PLEASE CLOSE THIS WINDOW and retry running the program. If the problem persists, please feel free to contact CORals for support at www.corals.no/kontakt.')
    exit()
