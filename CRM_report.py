from openpyxl import Workbook
import pandas
from openpyxl.utils.dataframe import dataframe_to_rows
import sys
import os
import numpy

#create the excel workbook

wb=Workbook()
ws1= wb.create_sheet("Sheet_1")
ws2= wb.create_sheet("Sheet_2")
ws3= wb.create_sheet("Sheet_3")
ws4= wb.create_sheet("Sheet_4")
ws5= wb.create_sheet("Sheet_5")
ws6= wb.create_sheet("Sheet_6")
ws7= wb.create_sheet("Sheet_7")


#add titles to the tabs 

ws1.title="Patient demographics"
ws2.title="Depth"
ws3.title="Variant-Calls"
ws4.title="Mutations and SNPS"
ws5.title="Screening Gaps"
ws6.title="Report"
ws7.title="NTC variant"


#add titles to give each sheet a structure

#Add titles to patient demographics tab
ws1['A1']='Date Received'
ws1['B1']='Leeds/Cardiff'
ws1['C1']='Lab number'
ws1['D1']='Notes'
ws1['E1']='Patient name'
ws1['F1']='Tumour %'
ws1['G1']='Qubit[DNA] ng/ul'
ws1['H1']='Dilution (ng/ul)'
ws1['I1']='NGS weeks'
ws1['J1']='Date set up'
ws1['K1']='Date of MiSeq run'
ws1['L1']='Library ng/ul(Qubit)'
ws1['M1']='Library nM'
ws1['N1']='Result'
ws1['O1']='Qubit conc. ng/ul'
ws1['P1']='Dilution (ng/ul)'
ws1['Q1']='Library ng/ul (Qubit)'
ws1['R1']='Library (nM)'


#Add titles to 1st depth table

ws2['A1']='CoverageCalculator'
ws2['A2']='Gene'
ws2['B2']='Full gene ROI'
ws2['C2']='Average depth'
ws2['D2']='Median depth'
ws2['E2']='Min read Depth'
ws2['F2']='NTC average depth'
ws2['G2']='%NTC'

ws2['A3']= 'NRAS_Ex4_c146'
ws2['B3']='c436_438'

ws2['A4']='NRAS_Ex4_c117'
ws2['B4']='c.349_351'

ws2['A5']='NRAS_Ex3_c61'
ws2['B5']='c.181_183'


ws2['A6']='NRAS_Ex2_c12_13'
ws2['B6']='c.34_39'

ws2['A8']='PIK3CA_Ex10_c542_546'
ws2['B8']='c.1624_1638'

ws2['A9']='PI3CA_Ex21_c1047_1049'
ws2['B9']='c.3139_3147'

ws2['A11']='BRAF_Ex15_c599_601'
ws2['B11']='c1795_c.1803'


ws2['A13']='KRAS_Ex4_c146'
ws2['B13']='c.436_438'

ws2['A14']='KRAS_Ex4_c117'
ws2['B14']='c.349_351'

ws2['A15']='KRAS_Ex3_c61'
ws2['B15']='c.181_183'

ws2['A16']='KRAS_Ex2_c12_13'
ws2['B16']='c.34_39'

ws2['A18']='TS53 Ex12'
ws2['B18']='c.1101-5_c.*5'

ws2['A19']='TP53 Ex11'
ws2['B19']='c.994-5_c.1100+5'

ws2['A20']='TP53 Ex9A & Ex9B'
ws2['B20']='c.993+191_c.993+333'

ws2['A21']='TP53 Ex9'
ws2['B21']='c.920_5_c.993+5'

ws2['A22']='TP53 Ex8'
ws2['B22']='c.783-5_c.919+5'

ws2['A23']='TP53 Ex7'
ws2['B23']='c.673-5_c.782+5'

ws2['A24']='TP53 Ex6'
ws2['B24']='c.560-5_c.672+5'

ws2['A25']='TP53 Ex5'
ws2['B25']='c.376-5_c.559+5'

ws2['A26']='TP53 Ex4'
ws2['B26']='c.97-5_c.96+5'

ws2['A27']='TP53 Ex3'
ws2['B27']='c.75_5_c.96+5'

ws2['A28']='TP53 Ex2'
ws2['B28']='c.-28-5_c.74+5'

#Add titles to second table

ws2['A31']= 'NRAS_Ex4_c146'
ws2['B31']='c436_438'

ws2['A32']='NRAS_Ex4_c117'
ws2['B32']='c.349_351'

ws2['A33']='NRAS_Ex3_c61'
ws2['B33']='c.181_183'


ws2['A34']='NRAS_Ex2_c12_13'
ws2['B34']='c.34_39'

ws2['A36']='PIK3CA_Ex10_c542_546'
ws2['B36']='c.1624_1638'

ws2['A37']='PIK3CA_Ex21_c1047_1049'
ws2['B37']='c.3139_3147'

ws2['A39']='BRAF_Ex15_c599_601'
ws2['B39']='c1795_c.1803'


ws2['A41']='KRAS_Ex4_c146'
ws2['B41']='c.436_438'

ws2['A42']='KRAS_Ex4_c117'
ws2['B42']='c.349_351'

ws2['A43']='KRAS_Ex3_c61'
ws2['B43']='c.181_183'

ws2['A44']='KRAS_Ex2_c12_13'
ws2['B44']='c.34_39'

ws2['A46']='TS53 Ex12'
ws2['B46']='c.1101-5_c.*5'

ws2['A47']='TP53 Ex11'
ws2['B47']='c.994-5_c.1100+5'

ws2['A48']='TP53 Ex9A & Ex9B'
ws2['B48']='c.993+191_c.993+333'

ws2['A49']='TP53 Ex9'
ws2['B49']='c.920_5_c.993+5'

ws2['A50']='TP53 Ex8'
ws2['B50']='c.783-5_c.919+5'

ws2['A51']='TP53 Ex7'
ws2['B51']='c.673-5_c.782+5'

ws2['A52']='TP53 Ex6'
ws2['B52']='c.560-5_c.672+5'

ws2['A53']='TP53 Ex5'
ws2['B53']='c.376-5_c.559+5'

ws2['A54']='TP53 Ex4'
ws2['B54']='c.97-5_c.96+5'

ws2['A55']='TP53 Ex3'
ws2['B55']='c.75_5_c.96+5'

ws2['A56']='TP53 Ex2'
ws2['B56']='c.-28-5_c.74+5'


#add titles to variant calls tab

ws3['E3']='NTC check1'
ws3['E6']='NTC check 2'

ws3['G3']='1st checker name & date'
ws3['G6']='2nd checker name & date'

ws3['K3']='GeneRead worksheet'
ws3['K6']='CRM panel'
ws3['A8']=" "


# Add titles to Mutations and SNPs tab
ws4['B2']='Gene'
ws4['C2']='Exon/Intron'
ws4['D2']='HGVS c'
ws4['E2']='HGVS p'
ws4['F2']='Allele Frequency'
ws4['G2']='Quality'
ws4['H2']='Depth'
ws4['I2']='Classification'
ws4['J2']='Transcript'
ws4['K2']='Variant'
ws4['L2']='Conclusion 1st checker'
ws4['M2']='Conclusion 2nd checker'


#Add titles to report tab

ws6['A4']='Lab number'
ws6['B4']='Patient name'
ws6['C4']='Tumour %'
ws6['D4']='Analysis'
ws6['E4']='Panel run'
ws6['F4']='Qubit[DNA] ng/ul'
ws6['G4']='Dilution (ng/ul)'
ws6['H4']='Analysed by:'

ws6['E7']='NGS wks'
ws6['F7']='Date set up'
ws6['G7']='Date of MiSeq run'
ws6['H7']='Library ng/ul(Qubit)'
ws6['I7']='Library nM'
ws6['J7']='NTC check 1'
ws6['K7']='NTC check 2'

ws6['A11']='Confirmed variant calls'      
ws6['A12']='Gene'
ws6['B12']='Exon'
ws6['C12']='Variant'
ws6['D12']='HGVS c.'
ws6['E12']='HGVS p.'
ws6['F12']='Allele frequency'
ws6['G12']='Conclusion 1st checker'
ws6['H12']='Conclusion 2nd checker'

ws6['A22']='Regions'
ws6['A23']='NRAS_c146(c.436_438)'
ws6['A24']='NRAS_c117(c.349_351)'
ws6['A25']='NRAS_c61(c.181_183)'
ws6['A26']='NRAS_c12_13(c.34_39)'
ws6['A27']='PIK3CA_c542_546(c.1624_1638)'
ws6['A28']='PIK3CA_c1047_1049(c.3139_3147)'
ws6['A29']='BRAF_c599_601(c.1795_c.1803)'
ws6['A30']='KRAS_c146(c.436_438)'
ws6['A31']='KRAS_c117(c.349_351)'
ws6['A32']='KRAS_c61(c.181_183)'
ws6['A33']='KRAS_c12_13(c.34_39)'

ws6['A21']='Percentage of bases in ROI covered to 500x'
ws6['B22']='% bases'

ws6['C23']='TP53 Ex11'
ws6['C24']='TP53 Ex10'
ws6['C25']='TP53 Ex9A & 9B'
ws6['C26']='TP53 Ex9'
ws6['C27']='TP53 Ex8'
ws6['C28']='TP53 Ex7'
ws6['C29']='TP53 Ex6'
ws6['C30']='TP53 Ex5'
ws6['C31']='TP53 Ex4'
ws6['C32']='TP53 Ex3'
ws6['C33']='TP53 Ex2'
ws6['C34']='TP53 overall'

ws6['D22']='% bases'
ws6['E22']='Gaps in hotspots ROI'
ws6['F34']='Comments'

detection_threshold=[]
variant_in_NTC=[]


def get_NTC_depth(path, referral, worksheet, runid):

    '''
    Open the NTC depth of coverage file and calculate the average depth between each set of coordinates 
    '''

    #Read the NTC depth of coverage file

    depth_of_coverage_NTC= pandas.read_csv(path+ "NTC-"+worksheet+"-"+referral+"/"+runid+"_NTC-"+worksheet+"-"+referral+"_DepthOfCoverage", sep="\t")
    depth_of_coverage_NTC[['locus_chrom', 'locus_position']]= depth_of_coverage_NTC['Locus'].str.split(':', expand=True)
    depth_of_coverage_NTC['locus_position'] = depth_of_coverage_NTC['locus_position'].astype(int)    


    #Calculate the average depth in the NTC between each set of coordinates
    average_NTC = []

    coordinates_list = [['1',115252202,115252204],
                    ['1',115252289, 115252291],
                    ['1',115256528, 115256530],
                    ['1',115258743 , 115258748],
                    ['3', 178936082, 178936096],
                    ['3', 178952084, 178952092],
                    ['7', 140453132, 140453140],
                    ['12', 25378560, 25378562],
                    ['12', 25378647, 25378649],
                    ['12', 25380275, 25380277],
                    ['12', 25398280,25398285],
                    ['17',7572922, 7573013],
                    ['17',7573922,7574038 ],
                    ['17',7576520, 7576662],
                    ['17',7576848,7576931],
                    ['17',7577014, 7577160],
                    ['17',7577494, 7577613],
                    ['17',7578172,7578294],
                    ['17',7578366, 7578559],
                    ['17',7579307, 7579595],
                    ['17',7579695, 7579726],
                    ['17',7579834, 7579945]]


    for coordinate in coordinates_list:
    
        coordinate_chrom = coordinate[0]
        coordinate_start = coordinate[1]
        coordinate_end = coordinate[2]
    
        #filter depth of coverage df
        filtered_df = depth_of_coverage_NTC[(depth_of_coverage_NTC['locus_chrom'] == coordinate_chrom) &
                                        (depth_of_coverage_NTC['locus_position'] >= coordinate_start) &
                                        (depth_of_coverage_NTC['locus_position'] <= coordinate_end)]
        
        # and append average to list
        average_NTC.append(filtered_df[filtered_df.columns[3]].mean())

    

    ws2['F3']=average_NTC[0]
    ws2['F4']=average_NTC[1]
    ws2['F5']=average_NTC[2]
    ws2['F6']=average_NTC[3]
    ws2['F8']=average_NTC[4]
    ws2['F9']=average_NTC[5]
    ws2['F11']=average_NTC[6]
    ws2['F13']=average_NTC[7]
    ws2['F14']=average_NTC[8]
    ws2['F15']=average_NTC[9]
    ws2['F16']=average_NTC[10]
    ws2['F18']=average_NTC[11]
    ws2['F19']=average_NTC[12]
    ws2['F20']=average_NTC[13]
    ws2['F21']=average_NTC[14]
    ws2['F22']=average_NTC[15]
    ws2['F23']=average_NTC[16]
    ws2['F24']=average_NTC[17]
    ws2['F25']=average_NTC[18]
    ws2['F26']=average_NTC[19]
    ws2['F27']=average_NTC[20]
    ws2['F28']=average_NTC[21]

    return(average_NTC)



def get_sample_depth(path, referral, sampleid, runid, Average_NTC):
   
    '''
    Open the sample depth of coverage file and calculate the average, minimum and median depth between each set of coordinates   
    '''
   
    #Read in the sample depth of coverage file

    depth_of_coverage_sample= pandas.read_csv(path+ sampleid+"/" +runid+"_"+sampleid+"_DepthOfCoverage", sep="\t")
    depth_of_coverage_sample[['locus_chromosome', 'locus_coordinates']]= depth_of_coverage_sample['Locus'].str.split(':', expand=True)
    depth_of_coverage_sample['locus_coordinates'] = depth_of_coverage_sample['locus_coordinates'].astype(int)


    #calculate the average, median and minimum depth between each set of coordinates

    Average=[]
    min=[]
    median=[]
 
    coordinates_list = [['1',115252202, 115252204],
                       ['1',115252289, 115252291], 
                       ['1',115256528, 115256530],
                       ['1',115258743, 115258748],
                       ['3',178936082, 178936096], 
                       ['3',178952084, 178952092], 
                       ['7',140453132, 140453140], 
                       ['12',25378560, 25378562], 
                       ['12',25378647, 25378649], 
                       ['12',25380275, 25380277],
                       ['12',25398280,25398285],
                       ['17',7572922, 7573013], 
                       ['17',7573922, 7574038],
                       ['17',7576520, 7576662], 
                       ['17',7576848, 7576931], 
                       ['17',7577014, 7577160], 
                       ['17',7577494, 7577613],
                       ['17',7578172, 7578294], 
                       ['17',7578366, 7578559], 
                       ['17',7579307, 7579595], 
                       ['17',7579695, 7579726],
                       ['17',7579834, 7579945]]

    for coordinate in coordinates_list:

        coordinate_chrom = coordinate[0]
        coordinate_start = coordinate[1]
        coordinate_end = coordinate[2]

        #filter depth of coverage df
        filtered_df = depth_of_coverage_sample[(depth_of_coverage_sample['locus_chromosome'] == coordinate_chrom) &
                                    (depth_of_coverage_sample['locus_coordinates'] >= coordinate_start) &
                                    (depth_of_coverage_sample['locus_coordinates'] <= coordinate_end)]

        # and append average, minimum and median to list
        Average.append(filtered_df[filtered_df.columns[3]].mean())    
        min.append(filtered_df[filtered_df.columns[3]].min())
        median.append(filtered_df[filtered_df.columns[3]].median())

    #add the sample averages to the excel workbook
    ws2['C3']=Average[0]
    ws2['C4']=Average[1]
    ws2['C5']=Average[2]
    ws2['C6']=Average[3]
    ws2['C8']=Average[4]
    ws2['C9']=Average[5]
    ws2['C11']=Average[6]
    ws2['C13']=Average[7]
    ws2['C14']=Average[8]
    ws2['C15']=Average[9]
    ws2['C16']=Average[10]
    ws2['C18']=Average[11]
    ws2['C19']=Average[12]
    ws2['C20']=Average[13]
    ws2['C21']=Average[14]
    ws2['C22']=Average[15]
    ws2['C23']=Average[16]
    ws2['C24']=Average[17]
    ws2['C25']=Average[18]
    ws2['C26']=Average[19]
    ws2['C27']=Average[20]
    ws2['C28']=Average[21]


    #add the minimum depths to the excel workbook
    ws2['E3']=min[0]
    ws2['E4']=min[1]
    ws2['E5']=min[2]
    ws2['E6']=min[3]
    ws2['E8']=min[4]
    ws2['E9']=min[5]
    ws2['E11']=min[6]
    ws2['E13']=min[7]
    ws2['E14']=min[8]
    ws2['E15']=min[9]
    ws2['E16']=min[10]
    ws2['E18']=min[11]
    ws2['E19']=min[12]
    ws2['E20']=min[13]
    ws2['E21']=min[14]
    ws2['E22']=min[15]
    ws2['E23']=min[16]
    ws2['E24']=min[17]
    ws2['E25']=min[18]
    ws2['E26']=min[19]
    ws2['E27']=min[20]
    ws2['E28']=min[21]
 
    #add the median depths to the excel workbook
    ws2['D3']=median[0]
    ws2['D4']=median[1]
    ws2['D5']=median[2]
    ws2['D6']=median[3]
    ws2['D8']=median[4]
    ws2['D9']=median[5]
    ws2['D11']=median[6]
    ws2['D13']=median[7]
    ws2['D14']=median[8]
    ws2['D15']=median[9]
    ws2['D16']=median[10]
    ws2['D18']=median[11]
    ws2['D19']=median[12]
    ws2['D20']=median[13]
    ws2['D21']=median[14]
    ws2['D22']=median[15]
    ws2['D23']=median[16]
    ws2['D24']=median[17]
    ws2['D25']=median[18]
    ws2['D26']=median[19]
    ws2['D27']=median[20]
    ws2['D28']=median[21]

    #Calculate %NTC and add it to the workbook

    ws2['G3']= (Average_NTC[0]/Average[0])*100
    ws2['G4']= (Average_NTC[1]/Average[1])*100
    ws2['G5']= (Average_NTC[2]/Average[2])*100
    ws2['G6']= (Average_NTC[3]/Average[3])*100
    ws2['G8']= (Average_NTC[4]/Average[4])*100
    ws2['G9']= (Average_NTC[5]/Average[5])*100
    ws2['G11']= (Average_NTC[6]/Average[6])*100
    ws2['G13']= (Average_NTC[7]/Average[7])*100
    ws2['G14']= (Average_NTC[8]/Average[8])*100
    ws2['G15']= (Average_NTC[9]/Average[9])*100
    ws2['G16']= (Average_NTC[10]/Average[10])*100
    ws2['G18']= (Average_NTC[11]/Average[11])*100
    ws2['G19']= (Average_NTC[12]/Average[12])*100
    ws2['G20']= (Average_NTC[13]/Average[13])*100
    ws2['G21']= (Average_NTC[14]/Average[14])*100
    ws2['G22']= (Average_NTC[15]/Average[15])*100
    ws2['G23']= (Average_NTC[16]/Average[16])*100
    ws2['G24']= (Average_NTC[17]/Average[17])*100
    ws2['G25']= (Average_NTC[18]/Average[18])*100
    ws2['G26']= (Average_NTC[19]/Average[19])*100
    ws2['G27']= (Average_NTC[20]/Average[20])*100
    ws2['G28']=(Average_NTC[21]/Average[21])*100

    return (depth_of_coverage_sample) 



def calculate_coverage_500x(depth_of_coverage_sample):
    
    '''
    Determine the percentage of bases with given coordinates that have a coverage of 500x
    '''


    coordinates_list = [['1',115252202,115252204],
                       ['1',115252289, 115252291],
                       ['1',115256528, 115256530],
                       ['1',115258743 , 115258748],
                       ['3', 178936082, 178936096],
                       ['3', 178952084, 178952092], 
                       ['7', 140453132, 140453140],
                       ['12', 25378560, 25378562],
                       ['12', 25378647, 25378649],
                       ['12', 25380275, 25380277],
                       ['12', 25398280,25398285], 
                       ['17',7572922, 7573013],
                       ['17',7573922,7574038 ],
                       ['17',7576520,  7576662], 
                       ['17',7576848,7576931],
                       ['17',7577014, 7577160],
                       ['17',7577494, 7577613],
                       ['17',7578172,7578294],
                       ['17',7578366, 7578559],
                       ['17',7579307, 7579595],
                       ['17',7579695, 7579726],
                       ['17',7579834, 7579945]]


    count=[]
    count_2=[]

        
    for coordinate in coordinates_list:
        coordinate_chrom = coordinate[0]
        coordinate_start = coordinate[1]
        coordinate_end = coordinate[2]


        #filter depth of coverage df to only include a certain region 
        
        filtered_df = depth_of_coverage_sample[(depth_of_coverage_sample['locus_chromosome'] == coordinate_chrom) &
                                    (depth_of_coverage_sample['locus_coordinates'] >= coordinate_start) &
                                    (depth_of_coverage_sample['locus_coordinates'] <= coordinate_end)]
        

        #count the number of bases within the region
        
        count.append(filtered_df.shape[0])

        
        #filter depth of coverage df to only include coordinates within a certain region that also have a depth above 500     
        
        filtered_df = depth_of_coverage_sample[(depth_of_coverage_sample['locus_chromosome'] == coordinate_chrom) &
                                    (depth_of_coverage_sample['locus_coordinates'] >= coordinate_start) &
                                    (depth_of_coverage_sample['locus_coordinates'] <= coordinate_end) & 
                                    (depth_of_coverage_sample['Depth_for_'+sampleid]>=500)]
        

        #count the number of bases within the given regions where the depth is greater than or equal to 500
       
        count_2.append(filtered_df.shape[0])        


    #Add the counts of the bases within the given regions to the excel workbook

    ws2['C31']=count_2[0]
    ws2['C32']=count_2[1]
    ws2['C33']=count_2[2]
    ws2['C34']=count_2[3]
    ws2['C36']=count_2[4]
    ws2['C37']=count_2[5]
    ws2['C39']=count_2[6]
    ws2['C41']=count_2[7]
    ws2['C42']=count_2[8]
    ws2['C43']=count_2[9]
    ws2['C44']=count_2[10]
    ws2['C46']=count_2[11]
    ws2['C47']=count_2[12]
    ws2['C48']=count_2[13]
    ws2['C49']=count_2[14]
    ws2['C50']=count_2[15]
    ws2['C51']=count_2[16]
    ws2['C52']=count_2[17]
    ws2['C53']=count_2[18]
    ws2['C54']=count_2[19]
    ws2['C55']=count_2[20]
    ws2['C56']=count_2[21]

    #Add the counts of the bases with a depth greater than or equal to 500 within the given regions to the excel workbook
 
    ws2['D31']=count[0]
    ws2['D32']=count[1]
    ws2['D33']=count[2]
    ws2['D34']=count[3]
    ws2['D36']=count[4]
    ws2['D37']=count[5]
    ws2['D39']=count[6]
    ws2['D41']=count[7]
    ws2['D42']=count[8]
    ws2['D43']=count[9]
    ws2['D44']=count[10]
    ws2['D46']=count[11]
    ws2['D47']=count[12]
    ws2['D48']=count[13]
    ws2['D49']=count[14]
    ws2['D50']=count[15]
    ws2['D51']=count[16]
    ws2['D52']=count[17]
    ws2['D53']=count[18]
    ws2['D54']=count[19]
    ws2['D55']=count[20]
    ws2['D56']=count[21]


    #Calculate percentage of bases that have a coverage of 500x and add to Depth tab

    ws2['E31']=(count_2[0]/count[0])
    ws2['E32']=(count_2[1]/count[1])
    ws2['E33']=(count_2[2]/count[2])
    ws2['E34']=(count_2[3]/count[3])
    ws2['E36']=(count_2[4]/count[4])
    ws2['E37']=(count_2[5]/count[5])
    ws2['E39']=(count_2[6]/count[6])
    ws2['E41']=(count_2[7]/count[7])
    ws2['E42']=(count_2[8]/count[8])
    ws2['E43']=(count_2[9]/count[9])
    ws2['E44']=(count_2[10]/count[10])
    ws2['E46']=(count_2[11]/count[11])
    ws2['E47']=(count_2[12]/count[12])
    ws2['E48']=(count_2[13]/count[13])
    ws2['E49']=(count_2[14]/count[14])
    ws2['E50']=(count_2[15]/count[15])
    ws2['E51']=(count_2[16]/count[16])
    ws2['E52']=(count_2[17]/count[17])
    ws2['E53']=(count_2[18]/count[18])
    ws2['E54']=(count_2[19]/count[19])
    ws2['E55']=(count_2[20]/count[20])
    ws2['E56']=(count_2[21]/count[21])


    #Calculate percentage of bases that	have a coverage	of 500x	and add	to Report tab
   
    ws6['B23']=(count_2[0]/count[0])*100
    ws6['B24']=(count_2[1]/count[1])*100
    ws6['B25']=(count_2[2]/count[2])*100
    ws6['B26']=(count_2[3]/count[3])*100
    ws6['B27']=(count_2[4]/count[4])*100
    ws6['B28']=(count_2[5]/count[5])*100
    ws6['B29']=(count_2[6]/count[6])*100
    ws6['B30']=(count_2[7]/count[7])*100
    ws6['B31']=(count_2[8]/count[8])*100
    ws6['B32']=(count_2[9]/count[9])*100
    ws6['B33']=(count_2[10]/count[10])*100
    ws6['D23']=(count_2[11]/count[11])*100
    ws6['D24']=(count_2[12]/count[12])*100
    ws6['D25']=(count_2[13]/count[13])*100
    ws6['D26']=(count_2[14]/count[14])*100
    ws6['D27']=(count_2[15]/count[15])*100
    ws6['D28']=(count_2[16]/count[16])*100
    ws6['D29']=(count_2[17]/count[17])*100
    ws6['D30']=(count_2[18]/count[18])*100
    ws6['D31']=(count_2[19]/count[19])*100
    ws6['D32']=(count_2[20]/count[20])*100
    ws6['D33']=(count_2[21]/count[21])*100


    #Calculate the percentage of bases with coverage 500x for the given regions in TP53

    ws2['C57']="=SUM('Depth'!C46:'Depth'!C56)"
    ws2['D57']="=SUM('Depth'!D46:'Depth'!D56)"
    ws2['E57']="='Depth'!C57/'Depth'!D57"
    ws6['D34']=("='Depth'!E57*100")



def get_variants(path, worksheet, referral, sampleid, runid):

    '''
    Fill the variant_calls and NTC variant tabs using the relevant files
    Determine if the variants in the sample are also present in the NTC
    If so, determine the level of NTC contamination of the variant in the sample
    Create variant alleles calls column by multiplying allele frequency by depth 
    '''    
   
    #Read the NTC variant report

    variant_report_NTC_2=pandas.read_csv(path+ "NTC-"+worksheet+"-"+referral+"/hotspot_variants/"+runid+"_NTC-"+worksheet+"-"+referral+"_"+ referral+"_VariantReport.txt", sep="\t")
    variant_report_NTC_3=variant_report_NTC_2[variant_report_NTC_2.PreferredTranscript!=False]
    variant_report_NTC_4= variant_report_NTC_3.iloc[:,[23,29,25,26,2,5,3,6,24,1]]


    #open relevant variants file for sample

    variant_report_2=pandas.read_csv(path+sampleid+"/hotspot_variants/"+runid+"_"+sampleid+"_"+referral+"_VariantReport.txt", sep="\t")
    ws6['A9']=runid+"_"+sampleid+"_"+referral+"_VariantReport.txt"
    variant_report_3=variant_report_2[variant_report_2.PreferredTranscript!=False]
    variant_report_4= variant_report_3.iloc[:,[23,29,25,26,2,5,3,6,24,1]]

    
    #Add Present in sample column to the NTC tab to show if the variants in the NTC are also present in the sample
    
    num_rows_NTC=variant_report_NTC_4.shape[0]
    num_rows_variant_report=variant_report_4.shape[0]
    variant_in_sample=[]
    row=0
    while (row<num_rows_NTC):
        row2=0
        present='NO'
        while (row2<num_rows_variant_report):
            if (variant_report_4.iloc[row2,9]==variant_report_NTC_4.iloc[row,9]):
                present='YES'
            row2=row2+1
        variant_in_sample.append(present)
        row=row+1    
    variant_report_NTC_4['Present in sample']=variant_in_sample


    #Add variant allele calls column (for each variant, allele frequency is multiplied by depth)

    variant_allele_calls=[]
    row=0
    num_rows_NTC=variant_report_NTC_4.shape[0]
    while (row<num_rows_NTC):
        variant_report_NTC_4.iloc[row,4]= variant_report_NTC_4.iloc[row,4].strip('%')
        variant_report_NTC_4.iloc[row,4]= float(variant_report_NTC_4.iloc[row,4])
        variant_report_NTC_4.iloc[row,4]= (variant_report_NTC_4.iloc[row,4])/100
        variant_report_NTC_4.iloc[row,6]=int(variant_report_NTC_4.iloc[row,6])
        allele_call= variant_report_NTC_4.iloc[row,4]*variant_report_NTC_4.iloc[row,6]
        variant_allele_calls.append(allele_call)
        row=row+1

    variant_report_NTC_4['Variant allele calls']=variant_allele_calls

    for row in dataframe_to_rows(variant_report_NTC_4, header=True, index=False):
        ws7.append(row)



    #create the extra table at the side of the variant_calls tab to determine if each variant is present in the NTC, and the level of NTC contamination

    row=0
    num_rows_variant_report=variant_report_4.shape[0]      

    while (row<num_rows_variant_report):
        variant_report_4.iloc[row,4]= variant_report_4.iloc[row,4].strip('%')
        variant_report_4.iloc[row,4]= float(variant_report_4.iloc[row,4])
        variant_report_4.iloc[row,4]= (variant_report_4.iloc[row,4])/100 
        variant_report_4.iloc[row,6]= int(variant_report_4.iloc[row,6])
        if (variant_report_4.iloc[row,6]<=500):
            value_2= variant_report_4.iloc[row,4]*variant_report_4.iloc[row,6]
            value_2=str(value_2)
        else:
            value_2=0
        detection_threshold.append(value_2)
        row2=0
        num_rows_NTC=variant_report_NTC_4.shape[0]
        present='NO'
        while (row2<num_rows_NTC):
            if (variant_report_4.iloc[row,9]==variant_report_NTC_4.iloc[row2,9]):
                present='YES'
            row2=row2+1
        variant_in_NTC.append(present)
        row=row+1

        
    #Add columns to the variant calls tab to enable the conclusion of each variant to be filled in

    variant_report_4["Conclusion 1st checker"]=""
    variant_report_4["QC"]=""
    variant_report_4["Conclusion 2nd checker"]=""
    variant_report_4["QC "]=""

    return (variant_report_4, variant_report_NTC_4)




def get_poly_artefacts(variant_report_4, variant_report_NTC_4):

    '''
    match each of the variants in the variant calls tab with those in "CRM_Poly and Artefact list.xlsx" to determine if they are known polys or artefacts 
    Create the second table in the variant calls tab to give the reason for each of the conclusions for the variant
    '''

    poly_artefact_dict={}
    poly_and_Artefact_list=pandas.read_excel("/home/transfer/pipelines/CRMworksheetCreator/CRM_poly_artefact_list.xlsx")
    poly_and_Artefact_list_2=pandas.DataFrame(poly_and_Artefact_list)

    num_rows_variant_report=variant_report_4.shape[0]
    num_rows_poly_artefact=poly_and_Artefact_list_2.shape[0]
    num_rows_NTC=variant_report_NTC_4.shape[0]

    row1=0
    while (row1<num_rows_variant_report):
        row2=0
        while(row2<num_rows_poly_artefact):
            if (poly_and_Artefact_list_2.iloc[row2,2]==variant_report_4.iloc[row1,2]):
                poly_artefact_dict[variant_report_4.iloc[row1,2]]= poly_and_Artefact_list_2.iloc[row2,2]
                variant_report_4.iloc[row1,10]= poly_and_Artefact_list_2.iloc[row2,6]
                variant_report_4.iloc[row1,12]= poly_and_Artefact_list_2.iloc[row2,6]
            row2=row2+1
        row1=row1+1    
    
    row3=0
    while (row3<num_rows_variant_report):
        for x in poly_artefact_dict:
            if (variant_report_4.iloc[row3,2]==x):
                if (variant_report_4.iloc[row3,10]=='Known artefact'):
                    variant_report_4.iloc[row3,11]=3
                    variant_report_4.iloc[row3,13]=3
                if (variant_report_4.iloc[row3,10]=='Known Poly'):
                    variant_report_4.iloc[row3,11]=1
                    variant_report_4.iloc[row3,13]=1
                if (variant_report_4.iloc[row3,10]=='WT'):
                    variant_report_4.iloc[row3,11]=3
                    variant_report_4.iloc[row3,13]=3
                if (variant_report_4.iloc[row3,10]=='Genuine'):
                    variant_report_4.iloc[row3,11]=1
                    variant_report_4.iloc[row3,13]=1
                if (variant_report_4.iloc[row3,10]=='SNP'):
                    variant_report_4.iloc[row3,11]=1
                    variant_report_4.iloc[row3,13]=1      
        row3=row3+1
     

    variant_report_4["Detection threshold based on depth"]=detection_threshold
    variant_report_4["Is variant present in NTC "]=variant_in_NTC
    variant_report_4["#of mutant reads in patient sample "]=""
    variant_report_4["#of mutant reads in NTC if present "]=""
    variant_report_4["Is the NTC contamination significant?"]=""


    #Determine if the level of NTC contamination is significant

    row=0
    while (row<num_rows_variant_report):
        if variant_report_4.iloc[row,15]=="YES":
            value2= variant_report_4.iloc[row,4]*variant_report_4.iloc[row,6]
            variant_report_4.iloc[row,16]=value2
        row2=0
        while (row2<num_rows_NTC): 
            if variant_report_4.iloc[row, 9]==variant_report_NTC_4.iloc[row2,9]:
                variant_report_4.iloc[row,17]=variant_report_NTC_4.iloc[row2,11]
                variant_report_4.iloc[row,18]= variant_report_4.iloc[row,17]/variant_report_4.iloc[row,16]
            row2=row2+1
        row=row+1        

    for row in dataframe_to_rows(variant_report_4, header=True, index=False):
        ws3.append(row)



    #create the second table in the variant calls tab

    variant_report_5= variant_report_4.iloc[:,[0,1,2]]
    variant_report_5['Comments/Notes/evidence:how conclusion was reached']=""

    row=0

    num_rows_variant_report=variant_report_4.shape[0]

    while (row<num_rows_variant_report):
        if (variant_report_4.iloc[row,10]=='Known artefact'):
            variant_report_5.iloc[row,3]='On artefact list'
        if (variant_report_4.iloc[row,10]=='Known Poly'):
            variant_report_5.iloc[row,3]='On Poly list'
        if (variant_report_4.iloc[row,10]=='WT'):
            variant_report_5.iloc[row,3]='SNP in Ref.Seq'     
        row=row+1
    

    ws3['A22']=" "

    #add dataframe to variant calls tab
    for row in dataframe_to_rows(variant_report_5, header=True, index=False):
        ws3.append(row)



def get_gaps(referral,path, sampleid, runid):
    
    '''
    open the relevant bed files and append these to the end of the screening gaps tab. If the bed files are empty, write 'no gaps'.
    '''

    hotspot_variants=referral
    #open the relevant bed files (ready for screening gaps tab)
    if (hotspot_variants=='GIST'):
        if((os.path.getsize(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_KIT_gaps.bed")==0) and (os.path.getsize(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_PDGFRA_gaps.bed")==0)):
            ws5['A1']= 'No gaps'
        if (os.path.getsize(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_KIT_gaps.bed")>0):
            bedfile_KIT=pandas.read_csv(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_KIT_gaps.bed", sep="\t")
            for row in dataframe_to_rows(bedfile_KIT, header=True, index=False):
                ws5.append(row)
        if (os.path.getsize(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_PDGFRA_gaps.bed")>0):
            bedfile_PDGFRA=pandas.read_csv(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_PDGFRA_gaps.bed", sep="\t")
            for row in dataframe_to_rows(bedfile_PDGFRA, header=True, index=False):
                ws5.append(row)
    elif (hotspot_variants=='FOCUS4'):
        if((os.path.getsize(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_BRAF_gaps.bed")==0) and (os.path.getsize(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_KRAS_gaps.bed")==0) and (os.path.getsize(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_NRAS_gaps.bed")==0) and (os.path.getsize(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_PIK3CA_gaps.bed")==0)and (os.path.getsize(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_TP53_gaps.bed")==0)):
            ws5['A1']= 'No gaps'
        if((os.path.getsize(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_BRAF_gaps.bed")>0)):
            bedfile_BRAF=pandas.read_csv(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_BRAF_gaps.bed", sep="\t")
            for row in dataframe_to_rows(bedfile_BRAF, header=True, index=False):
                ws5.append(row)
        if((os.path.getsize(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_KRAS_gaps.bed")>0)):
            bedfile_KRAS=pandas.read_csv(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_KRAS_gaps.bed", sep="\t")
            for row in dataframe_to_rows(bedfile_KRAS, header=True, index=False):
                ws5.append(row)
        if((os.path.getsize(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_NRAS_gaps.bed")>0)):
            bedfile_NRAS=pandas.read_csv(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_NRAS_gaps.bed", sep="\t")
            for row in dataframe_to_rows(bedfile_NRAS, header=True, index=False):
                ws5.append(row)
        if((os.path.getsize(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_PIK3CA_gaps.bed")>0)):
            bedfile_PIK3CA=pandas.read_csv(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_PIK3CA_gaps.bed", sep="\t")
            for row in dataframe_to_rows(bedfile_PIK3CA, header=True, index=False):
                ws5.append(row)
        if((os.path.getsize(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_TP53_gaps.bed")>0)):
            bedfile_TP53=pandas.read_csv(path+ sampleid+ "/hotspot_coverage/"+runid+"_"+sampleid+"_TP53_gaps.bed", sep="\t")
            for row in dataframe_to_rows(bedfile_TP53, header=True, index=False):
                ws5.append(row)


def add_excel_formulae():

    '''
    add excel formulae to the spreadsheets
    '''

    ws6['A13']= "='Mutations and SNPS'!B3"
    ws6['B13']= "='Mutations and SNPS'!C3"
    ws6['C13']= "='Mutations and SNPS'!K3"
    ws6['D13']= "='Mutations and SNPS'!D3"
    ws6['E13']= "='Mutations and SNPS'!E3"
    ws6['F13']= "='Mutations and SNPS'!F3"
    ws6['G13']= "='Mutations and SNPS'!L3"
    ws6['H13']= "='Mutations and SNPS'!M3"

    ws6['A14']= "='Mutations and SNPS'!B4"
    ws6['B14']= "='Mutations and SNPS'!C4"
    ws6['C14']= "='Mutations and SNPS'!K4"
    ws6['D14']= "='Mutations and SNPS'!D4"  
    ws6['E14']= "='Mutations and SNPS'!E4"
    ws6['F14']= "='Mutations and SNPS'!F4"
    ws6['G14']= "='Mutations and SNPS'!L4"
    ws6['H14']= "='Mutations and SNPS'!M4"

    ws6['A15']= "='Mutations and SNPS'!B5"
    ws6['B15']= "='Mutations and SNPS'!C5"
    ws6['C15']= "='Mutations and SNPS'!K5"
    ws6['D15']= "='Mutations and SNPS'!D5"
    ws6['E15']= "='Mutations and SNPS'!E5"
    ws6['F15']= "='Mutations and SNPS'!F5"
    ws6['G15']= "='Mutations and SNPS'!L5"
    ws6['H15']= "='Mutations and SNPS'!M5"

    ws6['A16']= "='Mutations and SNPS'!B6"
    ws6['B16']= "='Mutations and SNPS'!C6"
    ws6['C16']= "='Mutations and SNPS'!K6"
    ws6['D16']= "='Mutations and SNPS'!D6"
    ws6['E16']= "='Mutations and SNPS'!E6"
    ws6['F16']= "='Mutations and SNPS'!F6"
    ws6['G16']= "='Mutations and SNPS'!L6"
    ws6['H16']= "='Mutations and SNPS'!M6"

    ws6['A17']= "='Mutations and SNPS'!B7"
    ws6['B17']= "='Mutations and SNPS'!C7"
    ws6['C17']= "='Mutations and SNPS'!K7"
    ws6['D17']= "='Mutations and SNPS'!D7"
    ws6['E17']= "='Mutations and SNPS'!E7"
    ws6['F17']= "='Mutations and SNPS'!F7"
    ws6['G17']= "='Mutations and SNPS'!L7"
    ws6['H17']= "='Mutations and SNPS'!M7"

    ws6['A18']= "='Mutations and SNPS'!B8"
    ws6['B18']= "='Mutations and SNPS'!C8"
    ws6['C18']= "='Mutations and SNPS'!K8"
    ws6['D18']= "='Mutations and SNPS'!D8"
    ws6['E18']= "='Mutations and SNPS'!E8"
    ws6['F18']= "='Mutations and SNPS'!F8"
    ws6['G18']= "='Mutations and SNPS'!L8"
    ws6['H18']= "='Mutations and SNPS'!M8"

    ws6['A19']= "='Mutations and SNPS'!B9"
    ws6['B19']= "='Mutations and SNPS'!C9"
    ws6['C19']= "='Mutations and SNPS'!K9"
    ws6['D19']= "='Mutations and SNPS'!D9"
    ws6['E19']= "='Mutations and SNPS'!E9"
    ws6['F19']= "='Mutations and SNPS'!F9"
    ws6['G19']= "='Mutations and SNPS'!L9"
    ws6['H19']= "='Mutations and SNPS'!M9"


    ws6['E8']= "='Patient demographics'!I2" 
    ws6['F8']= "='Patient demographics'!J2"
    ws6['G8']= "='Patient demographics'!K2"
    ws6['H8']= "='Patient demographics'!L2"
    ws6['I8']= "='Patient demographics'!M2"
    ws6['J8']= "='Variant-Calls'!E4"
    ws6['K8']= "='Variant-Calls'!E7"


    ws6['A5']= sampleid
    ws6['B5']= "='Patient demographics'!E2"
    ws6['C5']= "='Patient demographics'!F2"
    ws6['D5']= "FOCUS4"
    ws6['E5']="CRM"
    ws6['F5']= "='Patient demographics'!G2"
    ws6['G5']= "='Patient demographics'!H2"
    ws6['H5']= "Checked by:"

    ws6['A1']= "=A5"
    ws6['B1']="FOCUS4 Patient Analysis sheet"

    ws3['O38']= "NTC check 1"
    ws3['O39']= "NTC check 2"


    ws2['O38']= "NTC check 1"
    ws2['O39']= "NTC check 2"

    ws3['E4']= "='Depth'!P38"
    ws3['K4']=worksheet

    ws6['E23']="='Screening Gaps'!D1"
    ws6['E24']="='Screening Gaps'!D2"
    ws6['E25']="='Screening Gaps'!D3"
    ws6['E26']="='Screening Gaps'!D4"
    ws6['E27']="='Screening Gaps'!D5"
    ws6['E28']="='Screening Gaps'!D6"
    ws6['E29']="='Screening Gaps'!D7"
    ws6['E30']="='Screening Gaps'!D8"
    ws6['E31']="='Screening Gaps'!D9"
    ws6['E32']="='Screening Gaps'!D10"
    ws6['E33']="='Screening Gaps'!D11"

    ws6['F23']="='Screening Gaps'!D12"
    ws6['F24']="='Screening Gaps'!D13"
    ws6['F25']="='Screening Gaps'!D14"
    ws6['F26']="='Screening Gaps'!D15"
    ws6['F27']="='Screening Gaps'!D16"
    ws6['F28']="='Screening Gaps'!D17"
    ws6['F29']="='Screening Gaps'!D18"
    ws6['F30']="='Screening Gaps'!D19"
    ws6['F31']="='Screening Gaps'!D20"
    ws6['F32']="='Screening Gaps'!D21"
    ws6['F33']="='Screening Gaps'!D22"



    wb.save(path+'/'+sampleid+'_'+referral+'_CRM.xlsx')




if __name__ == "__main__":


    #Insert information
    runid=sys.argv[1]
    sampleid=sys.argv[2]
    worksheet=sys.argv[3]
    referral=sys.argv[4]
    path= sys.argv[5]

    print(runid)
    print(sampleid)
    print(worksheet)
    print(referral)

    referrals_list=['FOCUS4', 'GIST']
    
    if referral in referrals_list:
        
        NTC_depth=get_NTC_depth(path, referral, worksheet, runid)
        
        sample_depth=get_sample_depth(path, referral, sampleid, runid,NTC_depth)

        calculate_coverage_500x(sample_depth)

        variant_report, variant_report_NTC=get_variants(path, worksheet, referral, sampleid, runid)

        get_poly_artefacts(variant_report, variant_report_NTC)

        get_gaps(referral,path, sampleid, runid)

        add_excel_formulae()
    else:
        print("referral not in referrals_list")
