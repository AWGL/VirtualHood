import unittest

from CRM_report_new_referrals import *

path="./tests_CRM/"

artefacts_path="./tests_CRM/"


class test_virtualhood(unittest.TestCase):


    def test_get_variantReport_NTC(self):
        self.assertEqual(len(get_variantReport_NTC("Colorectal", path, "NTC","test")),1)
        self.assertEqual(len(get_variantReport_NTC("Glioma", path, "NTC", "test")),7)
        self.assertEqual(len(get_variantReport_NTC("Lung", path, "NTC", "test")),5)
        self.assertEqual(len(get_variantReport_NTC("Melanoma", path, "NTC", "test")),0)
        self.assertEqual(len(get_variantReport_NTC("Thyroid", path, "NTC", "test")),3)


    def test_get_variant_report(self):
        self.assertEqual(len(get_variant_report("Colorectal", path, "tester", "test")),0)
        self.assertEqual(len(get_variant_report("Glioma", path, "tester", "test")),3)
        self.assertEqual(len(get_variant_report("Lung", path, "tester", "test")),5)
        self.assertEqual(len(get_variant_report("Melanoma", path, "tester", "test")),2)
        self.assertEqual(len(get_variant_report("Thyroid", path, "tester", "test")),4)




    def test_add_extra_columns_NTC_report(self):

        wb_Colorectal=Workbook()
        ws1_Colorectal= wb_Colorectal.create_sheet("Sheet_1")
        ws9_Colorectal= wb_Colorectal.create_sheet("Sheet_9")
        ws2_Colorectal= wb_Colorectal.create_sheet("Sheet_2")
        ws4_Colorectal= wb_Colorectal.create_sheet("Sheet_4")
        ws5_Colorectal= wb_Colorectal.create_sheet("Sheet_5")
        ws6_Colorectal= wb_Colorectal.create_sheet("Sheet_6")
        ws7_Colorectal= wb_Colorectal.create_sheet("Sheet_7")
        ws10_Colorectal= wb_Colorectal.create_sheet("Sheet_10")

        #name the tabs

        ws7_Colorectal.title="NTC variant"

        variant_report_NTC_Colorectal=get_variantReport_NTC("Colorectal", path, "NTC", "test")
        variant_report_Colorectal=get_variant_report("Colorectal", path, "tester", "test")


        variant_report_NTC, ws7=add_extra_columns_NTC_report(variant_report_NTC_Colorectal, variant_report_Colorectal, ws7_Colorectal, wb_Colorectal, path)
        self.assertEqual(ws7["A2"].value, "Gene1")
        self.assertEqual(ws7["B2"].value, "exon1")
        self.assertEqual(ws7["C2"].value, "HGVSv1")
        self.assertEqual(ws7["D2"].value, "HGVSp1")
        self.assertEqual(ws7["E2"].value, 1.0)
        self.assertEqual(ws7["F2"].value, "Genotype1")
        self.assertEqual(ws7["G2"].value, 5)
        self.assertEqual(ws7["H2"].value, "Quality1")
        self.assertEqual(ws7["I2"].value, "Transcript1")
        self.assertEqual(ws7["J2"].value, "variant1")
        self.assertEqual(ws7["K2"].value, "NO")
        self.assertEqual(ws7["L2"].value, 5.0)

        self.assertEqual(ws7["A3"].value, None)
        self.assertEqual(ws7["B3"].value, None)
        self.assertEqual(ws7["C3"].value, None)
        self.assertEqual(ws7["D3"].value, None)
        self.assertEqual(ws7["E3"].value, None)
        self.assertEqual(ws7["F3"].value, None)
        self.assertEqual(ws7["G3"].value, None)
        self.assertEqual(ws7["H3"].value, None)
        self.assertEqual(ws7["I3"].value, None)
        self.assertEqual(ws7["J3"].value, None)
        self.assertEqual(ws7["K3"].value, None)
        self.assertEqual(ws7["L3"].value, None)



#Glioma
        wb_Glioma=Workbook()
        ws1_Glioma= wb_Glioma.create_sheet("Sheet_1")
        ws9_Glioma= wb_Glioma.create_sheet("Sheet_9")
        ws2_Glioma= wb_Glioma.create_sheet("Sheet_2")
        ws4_Glioma= wb_Glioma.create_sheet("Sheet_4")
        ws5_Glioma= wb_Glioma.create_sheet("Sheet_5")
        ws6_Glioma= wb_Glioma.create_sheet("Sheet_6")
        ws7_Glioma= wb_Glioma.create_sheet("Sheet_7")
        ws10_Glioma= wb_Glioma.create_sheet("Sheet_10")

        #name the tabs
        ws1_Glioma.title="Patient demographics"
        ws2_Glioma.title="Variant_calls"
        ws4_Glioma.title="Mutations and SNPs"
        ws5_Glioma.title="hotspots.gaps"
        ws6_Glioma.title="Report"
        ws7_Glioma.title="NTC variant"
        ws9_Glioma.title="Subpanel NTC check"
        ws10_Glioma.title="Subpanel coverage"

        variant_report_NTC_Glioma=get_variantReport_NTC("Glioma", path, "NTC", "test")
        variant_report_Glioma=get_variant_report("Glioma", path, "tester", "test")


        variant_report_NTC, ws7=add_extra_columns_NTC_report(variant_report_NTC_Glioma, variant_report_Glioma, ws7_Glioma, wb_Glioma, path)
        self.assertEqual(ws7["A2"].value, "Gene1")
        self.assertEqual(ws7["B2"].value, "exon1")
        self.assertEqual(ws7["C2"].value, "HGVSv1")
        self.assertEqual(ws7["D2"].value, "HGVSp1")
        self.assertEqual(ws7["E2"].value, 1.0)
        self.assertEqual(ws7["F2"].value, "Genotype1")
        self.assertEqual(ws7["G2"].value, 5)
        self.assertEqual(ws7["H2"].value, "Quality1")
        self.assertEqual(ws7["I2"].value, "Transcript1")
        self.assertEqual(ws7["J2"].value, "variant1")
        self.assertEqual(ws7["K2"].value, "YES")
        self.assertEqual(ws7["L2"].value, 5.0)

        self.assertEqual(ws7["A3"].value, "Gene2")
        self.assertEqual(ws7["B3"].value, "exon2")
        self.assertEqual(ws7["C3"].value, "HGVSv2")
        self.assertEqual(ws7["D3"].value, "HGVSp2")
        self.assertEqual(ws7["E3"].value, 2.0)
        self.assertEqual(ws7["F3"].value, "Genotype2")
        self.assertEqual(ws7["G3"].value, 6)
        self.assertEqual(ws7["H3"].value, "Quality2")
        self.assertEqual(ws7["I3"].value, "Transcript2")
        self.assertEqual(ws7["J3"].value, "variant2")
        self.assertEqual(ws7["K3"].value, "YES")
        self.assertEqual(ws7["L3"].value, 12.0)

        self.assertEqual(ws7["A4"].value, "Gene3")
        self.assertEqual(ws7["B4"].value, "exon3")
        self.assertEqual(ws7["C4"].value, "HGVSv3")
        self.assertEqual(ws7["D4"].value, "HGVSp3")
        self.assertEqual(ws7["E4"].value, 3.0)
        self.assertEqual(ws7["F4"].value, "Genotype3")
        self.assertEqual(ws7["G4"].value, 7)
        self.assertEqual(ws7["H4"].value, "Quality3")
        self.assertEqual(ws7["I4"].value, "Transcript3")
        self.assertEqual(ws7["J4"].value, "variant3")
        self.assertEqual(ws7["K4"].value, "YES")
        self.assertEqual(ws7["L4"].value, 21.0)

        self.assertEqual(ws7["A5"].value, "Gene4")
        self.assertEqual(ws7["B5"].value, "exon4")
        self.assertEqual(ws7["C5"].value, "HGVSv4")
        self.assertEqual(ws7["D5"].value, "HGVSp4")
        self.assertEqual(ws7["E5"].value, 4.0)
        self.assertEqual(ws7["F5"].value, "Genotype4")
        self.assertEqual(ws7["G5"].value, 8)
        self.assertEqual(ws7["H5"].value, "Quality4")
        self.assertEqual(ws7["I5"].value, "Transcript4")
        self.assertEqual(ws7["J5"].value, "variant4")
        self.assertEqual(ws7["K5"].value, "NO")
        self.assertEqual(ws7["L5"].value, 32.0)

        self.assertEqual(ws7["A6"].value, "Gene5")
        self.assertEqual(ws7["B6"].value, "exon5")
        self.assertEqual(ws7["C6"].value, "HGVSv5")
        self.assertEqual(ws7["D6"].value, "HGVSp5")
        self.assertEqual(ws7["E6"].value, 5.0)
        self.assertEqual(ws7["F6"].value, "Genotype5")
        self.assertEqual(ws7["G6"].value, 9)
        self.assertEqual(ws7["H6"].value, "Quality5")
        self.assertEqual(ws7["I6"].value, "Transcript5")
        self.assertEqual(ws7["J6"].value, "variant5")
        self.assertEqual(ws7["K6"].value, "NO")
        self.assertEqual(ws7["L6"].value, 45.0)


        self.assertEqual(ws7["A7"].value, "Gene6")
        self.assertEqual(ws7["B7"].value, "exon6")
        self.assertEqual(ws7["C7"].value, "HGVSv6")
        self.assertEqual(ws7["D7"].value, "HGVSp6")
        self.assertEqual(ws7["E7"].value, 6.0)
        self.assertEqual(ws7["F7"].value, "Genotype6")
        self.assertEqual(ws7["G7"].value, 10)
        self.assertEqual(ws7["H7"].value, "Quality6")
        self.assertEqual(ws7["I7"].value, "Transcript6")
        self.assertEqual(ws7["J7"].value, "variant6")
        self.assertEqual(ws7["K7"].value, "NO")
        self.assertEqual(ws7["L7"].value, 60.0)


        self.assertEqual(ws7["A8"].value, "Gene7")
        self.assertEqual(ws7["B8"].value, "exon7")
        self.assertEqual(ws7["C8"].value, "HGVSv7")
        self.assertEqual(ws7["D8"].value, "HGVSp7")
        self.assertEqual(ws7["E8"].value, 7.0)
        self.assertEqual(ws7["F8"].value, "Genotype7")
        self.assertEqual(ws7["G8"].value, 11)
        self.assertEqual(ws7["H8"].value, "Quality7")
        self.assertEqual(ws7["I8"].value, "Transcript7")
        self.assertEqual(ws7["J8"].value, "variant7")
        self.assertEqual(ws7["K8"].value, "NO")
        self.assertEqual(ws7["L8"].value, 77.0)

        self.assertEqual(ws7["A9"].value, None)
        self.assertEqual(ws7["B9"].value, None)
        self.assertEqual(ws7["C9"].value, None)
        self.assertEqual(ws7["D9"].value, None)
        self.assertEqual(ws7["E9"].value, None)
        self.assertEqual(ws7["F9"].value, None)
        self.assertEqual(ws7["G9"].value, None)
        self.assertEqual(ws7["H9"].value, None)
        self.assertEqual(ws7["I9"].value, None)
        self.assertEqual(ws7["J9"].value, None)
        self.assertEqual(ws7["K9"].value, None)
        self.assertEqual(ws7["L9"].value, None)




#Lung
        wb_Lung=Workbook()
        ws1_Lung= wb_Lung.create_sheet("Sheet_1")
        ws9_Lung= wb_Lung.create_sheet("Sheet_9")
        ws2_Lung= wb_Lung.create_sheet("Sheet_2")
        ws4_Lung= wb_Lung.create_sheet("Sheet_4")
        ws5_Lung= wb_Lung.create_sheet("Sheet_5")
        ws6_Lung= wb_Lung.create_sheet("Sheet_6")
        ws7_Lung= wb_Lung.create_sheet("Sheet_7")
        ws10_Lung= wb_Lung.create_sheet("Sheet_10")

        #name the tabs
        ws1_Lung.title="Patient demographics"
        ws2_Lung.title="Variant_calls"
        ws4_Lung.title="Mutations and SNPs"
        ws5_Lung.title="hotspots.gaps"
        ws6_Lung.title="Report"
        ws7_Lung.title="NTC variant"
        ws9_Lung.title="Subpanel NTC check"
        ws10_Lung.title="Subpanel coverage"

        variant_report_NTC_Lung=get_variantReport_NTC("Lung", path, "NTC", "test")
        variant_report_Lung=get_variant_report("Lung", path, "tester", "test")


        variant_report_NTC, ws7=add_extra_columns_NTC_report(variant_report_NTC_Lung, variant_report_Lung, ws7_Lung, wb_Lung, path)
        self.assertEqual(ws7["A2"].value, "Gene1")
        self.assertEqual(ws7["B2"].value, "exon1")
        self.assertEqual(ws7["C2"].value, "HGVSv1")
        self.assertEqual(ws7["D2"].value, "HGVSp1")
        self.assertEqual(ws7["E2"].value, 1.0)
        self.assertEqual(ws7["F2"].value, "Genotype1")
        self.assertEqual(ws7["G2"].value, 5)
        self.assertEqual(ws7["H2"].value, "Quality1")
        self.assertEqual(ws7["I2"].value, "Transcript1")
        self.assertEqual(ws7["J2"].value, "variant1")
        self.assertEqual(ws7["K2"].value, "YES")
        self.assertEqual(ws7["L2"].value, 5.0)

        self.assertEqual(ws7["A3"].value, "Gene2")
        self.assertEqual(ws7["B3"].value, "exon2")
        self.assertEqual(ws7["C3"].value, "HGVSv2")
        self.assertEqual(ws7["D3"].value, "HGVSp2")
        self.assertEqual(ws7["E3"].value, 2.0)
        self.assertEqual(ws7["F3"].value, "Genotype2")
        self.assertEqual(ws7["G3"].value, 6)
        self.assertEqual(ws7["H3"].value, "Quality2")
        self.assertEqual(ws7["I3"].value, "Transcript2")
        self.assertEqual(ws7["J3"].value, "variant2")
        self.assertEqual(ws7["K3"].value, "YES")
        self.assertEqual(ws7["L3"].value, 12.0)

        self.assertEqual(ws7["A4"].value, "Gene3")
        self.assertEqual(ws7["B4"].value, "exon3")
        self.assertEqual(ws7["C4"].value, "HGVSv3")
        self.assertEqual(ws7["D4"].value, "HGVSp3")
        self.assertEqual(ws7["E4"].value, 3.0)
        self.assertEqual(ws7["F4"].value, "Genotype3")
        self.assertEqual(ws7["G4"].value, 7)
        self.assertEqual(ws7["H4"].value, "Quality3")
        self.assertEqual(ws7["I4"].value, "Transcript3")
        self.assertEqual(ws7["J4"].value, "variant3")
        self.assertEqual(ws7["K4"].value, "YES")
        self.assertEqual(ws7["L4"].value, 21.0)

        self.assertEqual(ws7["A5"].value, "Gene4")
        self.assertEqual(ws7["B5"].value, "exon4")
        self.assertEqual(ws7["C5"].value, "HGVSv4")
        self.assertEqual(ws7["D5"].value, "HGVSp4")
        self.assertEqual(ws7["E5"].value, 4.0)
        self.assertEqual(ws7["F5"].value, "Genotype4")
        self.assertEqual(ws7["G5"].value, 8)
        self.assertEqual(ws7["H5"].value, "Quality4")
        self.assertEqual(ws7["I5"].value, "Transcript4")
        self.assertEqual(ws7["J5"].value, "variant4")
        self.assertEqual(ws7["K5"].value, "YES")
        self.assertEqual(ws7["L5"].value, 32.0)

        self.assertEqual(ws7["A6"].value, "Gene5")
        self.assertEqual(ws7["B6"].value, "exon5")
        self.assertEqual(ws7["C6"].value, "HGVSv5")
        self.assertEqual(ws7["D6"].value, "HGVSp5")
        self.assertEqual(ws7["E6"].value, 5.0)
        self.assertEqual(ws7["F6"].value, "Genotype5")
        self.assertEqual(ws7["G6"].value, 9)
        self.assertEqual(ws7["H6"].value, "Quality5")
        self.assertEqual(ws7["I6"].value, "Transcript5")
        self.assertEqual(ws7["J6"].value, "variant5")
        self.assertEqual(ws7["K6"].value, "YES")
        self.assertEqual(ws7["L6"].value, 45.0)

        self.assertEqual(ws7["A7"].value, None)
        self.assertEqual(ws7["B7"].value, None)
        self.assertEqual(ws7["C7"].value, None)
        self.assertEqual(ws7["D7"].value, None)
        self.assertEqual(ws7["E7"].value, None)
        self.assertEqual(ws7["F7"].value, None)
        self.assertEqual(ws7["G7"].value, None)
        self.assertEqual(ws7["H7"].value, None)
        self.assertEqual(ws7["I7"].value, None)
        self.assertEqual(ws7["J7"].value, None)
        self.assertEqual(ws7["K7"].value, None)
        self.assertEqual(ws7["L7"].value, None)



#Melanoma
        wb_Melanoma=Workbook()
        ws1_Melanoma= wb_Melanoma.create_sheet("Sheet_1")
        ws9_Melanoma= wb_Melanoma.create_sheet("Sheet_9")
        ws2_Melanoma= wb_Melanoma.create_sheet("Sheet_2")
        ws4_Melanoma= wb_Melanoma.create_sheet("Sheet_4")
        ws5_Melanoma= wb_Melanoma.create_sheet("Sheet_5")
        ws6_Melanoma= wb_Melanoma.create_sheet("Sheet_6")
        ws7_Melanoma= wb_Melanoma.create_sheet("Sheet_7")
        ws10_Melanoma= wb_Melanoma.create_sheet("Sheet_10")

        #name the tabs
        ws1_Melanoma.title="Patient demographics"
        ws2_Melanoma.title="Variant_calls"
        ws4_Melanoma.title="Mutations and SNPs"
        ws5_Melanoma.title="hotspots.gaps"
        ws6_Melanoma.title="Report"
        ws7_Melanoma.title="NTC variant"
        ws9_Melanoma.title="Subpanel NTC check"
        ws10_Melanoma.title="Subpanel coverage"

        variant_report_NTC_Melanoma=get_variantReport_NTC("Melanoma", path, "NTC", "test")
        variant_report_Melanoma=get_variant_report("Melanoma", path, "tester", "test")


        variant_report_NTC, ws7=add_extra_columns_NTC_report(variant_report_NTC_Melanoma, variant_report_Melanoma, ws7_Melanoma, wb_Melanoma, path)
        self.assertEqual(ws7["A2"].value, None)
        self.assertEqual(ws7["B2"].value, None)
        self.assertEqual(ws7["C2"].value, None)
        self.assertEqual(ws7["D2"].value, None)
        self.assertEqual(ws7["E2"].value, None)
        self.assertEqual(ws7["F2"].value, None)
        self.assertEqual(ws7["G2"].value, None)
        self.assertEqual(ws7["H2"].value, None)
        self.assertEqual(ws7["I2"].value, None)
        self.assertEqual(ws7["J2"].value, None)
        self.assertEqual(ws7["K2"].value, None)
        self.assertEqual(ws7["L2"].value, None)




#Thyroid
        wb_Thyroid=Workbook()
        ws1_Thyroid= wb_Thyroid.create_sheet("Sheet_1")
        ws9_Thyroid= wb_Thyroid.create_sheet("Sheet_9")
        ws2_Thyroid= wb_Thyroid.create_sheet("Sheet_2")
        ws4_Thyroid= wb_Thyroid.create_sheet("Sheet_4")
        ws5_Thyroid= wb_Thyroid.create_sheet("Sheet_5")
        ws6_Thyroid= wb_Thyroid.create_sheet("Sheet_6")
        ws7_Thyroid= wb_Thyroid.create_sheet("Sheet_7")
        ws10_Thyroid= wb_Thyroid.create_sheet("Sheet_10")

        #name the tabs
        ws1_Thyroid.title="Patient demographics"
        ws2_Thyroid.title="Variant_calls"
        ws4_Thyroid.title="Mutations and SNPs"
        ws5_Thyroid.title="hotspots.gaps"
        ws6_Thyroid.title="Report"
        ws7_Thyroid.title="NTC variant"
        ws9_Thyroid.title="Subpanel NTC check"
        ws10_Thyroid.title="Subpanel coverage"

        variant_report_NTC_Thyroid=get_variantReport_NTC("Thyroid", path, "NTC", "test")
        variant_report_Thyroid=get_variant_report("Thyroid", path, "tester", "test")


        variant_report_NTC, ws7=add_extra_columns_NTC_report(variant_report_NTC_Thyroid, variant_report_Thyroid, ws7_Thyroid, wb_Thyroid, path)
        self.assertEqual(ws7["A2"].value, "Gene1")
        self.assertEqual(ws7["B2"].value, "exon1")
        self.assertEqual(ws7["C2"].value, "HGVSv1")
        self.assertEqual(ws7["D2"].value, "HGVSp1")
        self.assertEqual(ws7["E2"].value, 1.0)
        self.assertEqual(ws7["F2"].value, "Genotype1")
        self.assertEqual(ws7["G2"].value, 5)
        self.assertEqual(ws7["H2"].value, "Quality1")
        self.assertEqual(ws7["I2"].value, "Transcript1")
        self.assertEqual(ws7["J2"].value, "variant1")
        self.assertEqual(ws7["K2"].value, "YES")
        self.assertEqual(ws7["L2"].value, 5.0)

        self.assertEqual(ws7["A3"].value, "Gene2")
        self.assertEqual(ws7["B3"].value, "exon2")
        self.assertEqual(ws7["C3"].value, "HGVSv2")
        self.assertEqual(ws7["D3"].value, "HGVSp2")
        self.assertEqual(ws7["E3"].value, 2.0)
        self.assertEqual(ws7["F3"].value, "Genotype2")
        self.assertEqual(ws7["G3"].value, 6)
        self.assertEqual(ws7["H3"].value, "Quality2")
        self.assertEqual(ws7["I3"].value, "Transcript2")
        self.assertEqual(ws7["J3"].value, "variant2")
        self.assertEqual(ws7["K3"].value, "YES")
        self.assertEqual(ws7["L3"].value, 12.0)

        self.assertEqual(ws7["A4"].value, "Gene3")
        self.assertEqual(ws7["B4"].value, "exon3")
        self.assertEqual(ws7["C4"].value, "HGVSv3")
        self.assertEqual(ws7["D4"].value, "HGVSp3")
        self.assertEqual(ws7["E4"].value, 3.0)
        self.assertEqual(ws7["F4"].value, "Genotype3")
        self.assertEqual(ws7["G4"].value, 7)
        self.assertEqual(ws7["H4"].value, "Quality3")
        self.assertEqual(ws7["I4"].value, "Transcript3")
        self.assertEqual(ws7["J4"].value, "variant3")
        self.assertEqual(ws7["K4"].value, "YES")
        self.assertEqual(ws7["L4"].value, 21.0)

        self.assertEqual(ws7["A5"].value, None)
        self.assertEqual(ws7["B5"].value, None)
        self.assertEqual(ws7["C5"].value, None)
        self.assertEqual(ws7["D5"].value, None)
        self.assertEqual(ws7["E5"].value, None)
        self.assertEqual(ws7["F5"].value, None)
        self.assertEqual(ws7["G5"].value, None)
        self.assertEqual(ws7["H5"].value, None)
        self.assertEqual(ws7["I5"].value, None)
        self.assertEqual(ws7["J5"].value, None)
        self.assertEqual(ws7["K5"].value, None)
        self.assertEqual(ws7["L5"].value, None)



    def test_expand_variant_report(self):


#Colorectal
        variant_report_NTC_Colorectal=get_variantReport_NTC("Colorectal", path, "NTC", "test")
        variant_report_Colorectal=get_variant_report("Colorectal", path, "tester", "test")

        variant_report_Colorectal=expand_variant_report(variant_report_Colorectal, variant_report_NTC_Colorectal)

        self.assertEqual(len(variant_report_Colorectal),0)





#Glioma
        variant_report_NTC_Glioma=get_variantReport_NTC("Glioma", path, "NTC", "test")
        variant_report_Glioma=get_variant_report("Glioma", path, "tester", "test")

        variant_report_Glioma=expand_variant_report(variant_report_Glioma, variant_report_NTC_Glioma)


        self.assertEqual(variant_report_Glioma.iloc[0,0], "Gene1")
        self.assertEqual(variant_report_Glioma.iloc[0,1], "exon1")
        self.assertEqual(variant_report_Glioma.iloc[0,2], "HGVSv1")
        self.assertEqual(variant_report_Glioma.iloc[0,3], "HGVSp1" )
        self.assertEqual(variant_report_Glioma.iloc[0,4], 1.0)
        self.assertEqual(variant_report_Glioma.iloc[0,5], "Genotype1")
        self.assertEqual(variant_report_Glioma.iloc[0,6], 5)
        self.assertEqual(variant_report_Glioma.iloc[0,7], "Quality1")
        self.assertEqual(variant_report_Glioma.iloc[0,8], "Transcript1")
        self.assertEqual(variant_report_Glioma.iloc[0,9], "variant1")
        self.assertEqual(variant_report_Glioma.iloc[0,10], "")
        self.assertEqual(variant_report_Glioma.iloc[0,11], "")
        self.assertEqual(variant_report_Glioma.iloc[0,12], "")



        self.assertEqual(variant_report_Glioma.iloc[1,0], "Gene2")
        self.assertEqual(variant_report_Glioma.iloc[1,1], "exon2")
        self.assertEqual(variant_report_Glioma.iloc[1,2], "HGVSv2")
        self.assertEqual(variant_report_Glioma.iloc[1,3], "HGVSp2" )
        self.assertEqual(variant_report_Glioma.iloc[1,4], 2.0)
        self.assertEqual(variant_report_Glioma.iloc[1,5], "Genotype2")
        self.assertEqual(variant_report_Glioma.iloc[1,6], 6)
        self.assertEqual(variant_report_Glioma.iloc[1,7], "Quality2")
        self.assertEqual(variant_report_Glioma.iloc[1,8], "Transcript2")
        self.assertEqual(variant_report_Glioma.iloc[1,9], "variant2")
        self.assertEqual(variant_report_Glioma.iloc[1,10], "")
        self.assertEqual(variant_report_Glioma.iloc[1,11], "")
        self.assertEqual(variant_report_Glioma.iloc[1,12], "")


        self.assertEqual(variant_report_Glioma.iloc[2,0], "Gene3")
        self.assertEqual(variant_report_Glioma.iloc[2,1], "exon3")
        self.assertEqual(variant_report_Glioma.iloc[2,2], "HGVSv3")
        self.assertEqual(variant_report_Glioma.iloc[2,3], "HGVSp3" )
        self.assertEqual(variant_report_Glioma.iloc[2,4], 3.0)
        self.assertEqual(variant_report_Glioma.iloc[2,5], "Genotype3")
        self.assertEqual(variant_report_Glioma.iloc[2,6], 7)
        self.assertEqual(variant_report_Glioma.iloc[2,7], "Quality3")
        self.assertEqual(variant_report_Glioma.iloc[2,8], "Transcript3")
        self.assertEqual(variant_report_Glioma.iloc[2,9], "variant3")
        self.assertEqual(variant_report_Glioma.iloc[2,10], "")
        self.assertEqual(variant_report_Glioma.iloc[2,11], "")
        self.assertEqual(variant_report_Glioma.iloc[2,12], "")




#Lung
        variant_report_NTC_Lung=get_variantReport_NTC("Lung", path, "NTC", "test")
        variant_report_Lung=get_variant_report("Lung", path, "tester", "test")

        variant_report_Lung=expand_variant_report(variant_report_Lung, variant_report_NTC_Lung)


        self.assertEqual(variant_report_Lung.iloc[0,0], "Gene1")
        self.assertEqual(variant_report_Lung.iloc[0,1], "exon1")
        self.assertEqual(variant_report_Lung.iloc[0,2], "HGVSv1")
        self.assertEqual(variant_report_Lung.iloc[0,3], "HGVSp1" )
        self.assertEqual(variant_report_Lung.iloc[0,4], 1.0)
        self.assertEqual(variant_report_Lung.iloc[0,5], "Genotype1")
        self.assertEqual(variant_report_Lung.iloc[0,6], 5)
        self.assertEqual(variant_report_Lung.iloc[0,7], "Quality1")
        self.assertEqual(variant_report_Lung.iloc[0,8], "Transcript1")
        self.assertEqual(variant_report_Lung.iloc[0,9], "variant1")
        self.assertEqual(variant_report_Lung.iloc[0,10], "")
        self.assertEqual(variant_report_Lung.iloc[0,11], "")
        self.assertEqual(variant_report_Lung.iloc[0,12], "")



        self.assertEqual(variant_report_Lung.iloc[1,0], "Gene2")
        self.assertEqual(variant_report_Lung.iloc[1,1], "exon2")
        self.assertEqual(variant_report_Lung.iloc[1,2], "HGVSv2")
        self.assertEqual(variant_report_Lung.iloc[1,3], "HGVSp2" )
        self.assertEqual(variant_report_Lung.iloc[1,4], 2.0)
        self.assertEqual(variant_report_Lung.iloc[1,5], "Genotype2")
        self.assertEqual(variant_report_Lung.iloc[1,6], 6)
        self.assertEqual(variant_report_Lung.iloc[1,7], "Quality2")
        self.assertEqual(variant_report_Lung.iloc[1,8], "Transcript2")
        self.assertEqual(variant_report_Lung.iloc[1,9], "variant2")
        self.assertEqual(variant_report_Lung.iloc[1,10], "")
        self.assertEqual(variant_report_Lung.iloc[1,11], "")
        self.assertEqual(variant_report_Lung.iloc[1,12], "")


        self.assertEqual(variant_report_Lung.iloc[2,0], "Gene3")
        self.assertEqual(variant_report_Lung.iloc[2,1], "exon3")
        self.assertEqual(variant_report_Lung.iloc[2,2], "HGVSv3")
        self.assertEqual(variant_report_Lung.iloc[2,3], "HGVSp3" )
        self.assertEqual(variant_report_Lung.iloc[2,4], 3.0)
        self.assertEqual(variant_report_Lung.iloc[2,5], "Genotype3")
        self.assertEqual(variant_report_Lung.iloc[2,6], 7)
        self.assertEqual(variant_report_Lung.iloc[2,7], "Quality3")
        self.assertEqual(variant_report_Lung.iloc[2,8], "Transcript3")
        self.assertEqual(variant_report_Lung.iloc[2,9], "variant3")
        self.assertEqual(variant_report_Lung.iloc[2,10], "")
        self.assertEqual(variant_report_Lung.iloc[2,11], "")
        self.assertEqual(variant_report_Lung.iloc[2,12], "")

        self.assertEqual(variant_report_Lung.iloc[3,0], "Gene4")
        self.assertEqual(variant_report_Lung.iloc[3,1], "exon4")
        self.assertEqual(variant_report_Lung.iloc[3,2], "HGVSv4")
        self.assertEqual(variant_report_Lung.iloc[3,3], "HGVSp4" )
        self.assertEqual(variant_report_Lung.iloc[3,4], 4.0)
        self.assertEqual(variant_report_Lung.iloc[3,5], "Genotype4")
        self.assertEqual(variant_report_Lung.iloc[3,6], 8)
        self.assertEqual(variant_report_Lung.iloc[3,7], "Quality4")
        self.assertEqual(variant_report_Lung.iloc[3,8], "Transcript4")
        self.assertEqual(variant_report_Lung.iloc[3,9], "variant4")
        self.assertEqual(variant_report_Lung.iloc[3,10], "")
        self.assertEqual(variant_report_Lung.iloc[3,11], "")
        self.assertEqual(variant_report_Lung.iloc[3,12], "")

        self.assertEqual(variant_report_Lung.iloc[4,0], "Gene5")
        self.assertEqual(variant_report_Lung.iloc[4,1], "exon5")
        self.assertEqual(variant_report_Lung.iloc[4,2], "HGVSv5")
        self.assertEqual(variant_report_Lung.iloc[4,3], "HGVSp5" )
        self.assertEqual(variant_report_Lung.iloc[4,4], 5.0)
        self.assertEqual(variant_report_Lung.iloc[4,5], "Genotype5")
        self.assertEqual(variant_report_Lung.iloc[4,6], 9)
        self.assertEqual(variant_report_Lung.iloc[4,7], "Quality5")
        self.assertEqual(variant_report_Lung.iloc[4,8], "Transcript5")
        self.assertEqual(variant_report_Lung.iloc[4,9], "variant5")
        self.assertEqual(variant_report_Lung.iloc[4,10], "")
        self.assertEqual(variant_report_Lung.iloc[4,11], "")
        self.assertEqual(variant_report_Lung.iloc[4,12], "")



#Melanoma
        variant_report_NTC_Melanoma=get_variantReport_NTC("Melanoma", path, "NTC", "test")
        variant_report_Melanoma=get_variant_report("Melanoma", path, "tester", "test")

        variant_report_Melanoma=expand_variant_report(variant_report_Melanoma, variant_report_NTC_Melanoma)


        self.assertEqual(variant_report_Melanoma.iloc[0,0], "Gene1")
        self.assertEqual(variant_report_Melanoma.iloc[0,1], "exon1")
        self.assertEqual(variant_report_Melanoma.iloc[0,2], "HGVSv1")
        self.assertEqual(variant_report_Melanoma.iloc[0,3], "HGVSp1" )
        self.assertEqual(variant_report_Melanoma.iloc[0,4], 1.0)
        self.assertEqual(variant_report_Melanoma.iloc[0,5], "Genotype1")
        self.assertEqual(variant_report_Melanoma.iloc[0,6], 5)
        self.assertEqual(variant_report_Melanoma.iloc[0,7], "Quality1")
        self.assertEqual(variant_report_Melanoma.iloc[0,8], "Transcript1")
        self.assertEqual(variant_report_Melanoma.iloc[0,9], "variant1")
        self.assertEqual(variant_report_Melanoma.iloc[0,10], "")
        self.assertEqual(variant_report_Melanoma.iloc[0,11], "")
        self.assertEqual(variant_report_Melanoma.iloc[0,12], "")



        self.assertEqual(variant_report_Melanoma.iloc[1,0], "Gene2")
        self.assertEqual(variant_report_Melanoma.iloc[1,1], "exon2")
        self.assertEqual(variant_report_Melanoma.iloc[1,2], "HGVSv2")
        self.assertEqual(variant_report_Melanoma.iloc[1,3], "HGVSp2" )
        self.assertEqual(variant_report_Melanoma.iloc[1,4], 2.0)
        self.assertEqual(variant_report_Melanoma.iloc[1,5], "Genotype2")
        self.assertEqual(variant_report_Melanoma.iloc[1,6], 6)
        self.assertEqual(variant_report_Melanoma.iloc[1,7], "Quality2")
        self.assertEqual(variant_report_Melanoma.iloc[1,8], "Transcript2")
        self.assertEqual(variant_report_Melanoma.iloc[1,9], "variant2")
        self.assertEqual(variant_report_Melanoma.iloc[1,10], "")
        self.assertEqual(variant_report_Melanoma.iloc[1,11], "")
        self.assertEqual(variant_report_Melanoma.iloc[1,12], "")





#Thyroid
        variant_report_NTC_Thyroid=get_variantReport_NTC("Thyroid", path, "NTC", "test")
        variant_report_Thyroid=get_variant_report("Thyroid", path, "tester", "test")

        variant_report_Thyroid=expand_variant_report(variant_report_Thyroid, variant_report_NTC_Thyroid)


        self.assertEqual(variant_report_Thyroid.iloc[0,0], "Gene3")
        self.assertEqual(variant_report_Thyroid.iloc[0,1], "exon3")
        self.assertEqual(variant_report_Thyroid.iloc[0,2], "HGVSv3")
        self.assertEqual(variant_report_Thyroid.iloc[0,3], "HGVSp3" )
        self.assertEqual(variant_report_Thyroid.iloc[0,4], 3.0)
        self.assertEqual(variant_report_Thyroid.iloc[0,5], "Genotype3")
        self.assertEqual(variant_report_Thyroid.iloc[0,6], 5)
        self.assertEqual(variant_report_Thyroid.iloc[0,7], "Quality3")
        self.assertEqual(variant_report_Thyroid.iloc[0,8], "Transcript3")
        self.assertEqual(variant_report_Thyroid.iloc[0,9], "variant3")
        self.assertEqual(variant_report_Thyroid.iloc[0,10], "")
        self.assertEqual(variant_report_Thyroid.iloc[0,11], "")
        self.assertEqual(variant_report_Thyroid.iloc[0,12], "")



        self.assertEqual(variant_report_Thyroid.iloc[1,0], "Gene2")
        self.assertEqual(variant_report_Thyroid.iloc[1,1], "exon2")
        self.assertEqual(variant_report_Thyroid.iloc[1,2], "HGVSv2")
        self.assertEqual(variant_report_Thyroid.iloc[1,3], "HGVSp2" )
        self.assertEqual(variant_report_Thyroid.iloc[1,4], 2.0)
        self.assertEqual(variant_report_Thyroid.iloc[1,5], "Genotype2")
        self.assertEqual(variant_report_Thyroid.iloc[1,6], 6)
        self.assertEqual(variant_report_Thyroid.iloc[1,7], "Quality2")
        self.assertEqual(variant_report_Thyroid.iloc[1,8], "Transcript2")
        self.assertEqual(variant_report_Thyroid.iloc[1,9], "variant2")
        self.assertEqual(variant_report_Thyroid.iloc[1,10], "")
        self.assertEqual(variant_report_Thyroid.iloc[1,11], "")
        self.assertEqual(variant_report_Thyroid.iloc[1,12], "")


        self.assertEqual(variant_report_Thyroid.iloc[2,0], "Gene1")
        self.assertEqual(variant_report_Thyroid.iloc[2,1], "exon1")
        self.assertEqual(variant_report_Thyroid.iloc[2,2], "HGVSv1")
        self.assertEqual(variant_report_Thyroid.iloc[2,3], "HGVSp1" )
        self.assertEqual(variant_report_Thyroid.iloc[2,4], 1.0)
        self.assertEqual(variant_report_Thyroid.iloc[2,5], "Genotype1")
        self.assertEqual(variant_report_Thyroid.iloc[2,6], 7)
        self.assertEqual(variant_report_Thyroid.iloc[2,7], "Quality1")
        self.assertEqual(variant_report_Thyroid.iloc[2,8], "Transcript1")
        self.assertEqual(variant_report_Thyroid.iloc[2,9], "variant1")
        self.assertEqual(variant_report_Thyroid.iloc[2,10], "")
        self.assertEqual(variant_report_Thyroid.iloc[2,11], "")
        self.assertEqual(variant_report_Thyroid.iloc[2,12], "")

        self.assertEqual(variant_report_Thyroid.iloc[3,0], "Gene7")
        self.assertEqual(variant_report_Thyroid.iloc[3,1], "exon7")
        self.assertEqual(variant_report_Thyroid.iloc[3,2], "HGVSv7")
        self.assertEqual(variant_report_Thyroid.iloc[3,3], "HGVSp7" )
        self.assertEqual(variant_report_Thyroid.iloc[3,4], 7.0)
        self.assertEqual(variant_report_Thyroid.iloc[3,5], "Genotype7")
        self.assertEqual(variant_report_Thyroid.iloc[3,6], 11)
        self.assertEqual(variant_report_Thyroid.iloc[3,7], "Quality7")
        self.assertEqual(variant_report_Thyroid.iloc[3,8], "Transcript7")
        self.assertEqual(variant_report_Thyroid.iloc[3,9], "variant7")
        self.assertEqual(variant_report_Thyroid.iloc[3,10], "")
        self.assertEqual(variant_report_Thyroid.iloc[3,11], "")
        self.assertEqual(variant_report_Thyroid.iloc[3,12], "")





    def test_get_gaps_file(self):

	#Colorectal


        wb_Colorectal=Workbook()
        ws1_Colorectal= wb_Colorectal.create_sheet("Sheet_1")
        ws9_Colorectal= wb_Colorectal.create_sheet("Sheet_9")
        ws2_Colorectal= wb_Colorectal.create_sheet("Sheet_2")
        ws4_Colorectal= wb_Colorectal.create_sheet("Sheet_4")
        ws5_Colorectal= wb_Colorectal.create_sheet("Sheet_5")
        ws6_Colorectal= wb_Colorectal.create_sheet("Sheet_6")
        ws7_Colorectal= wb_Colorectal.create_sheet("Sheet_7")
        ws10_Colorectal=wb_Colorectal.create_sheet("Sheet_10")

        #name the tabs
        ws1_Colorectal.title="Patient demographics"
        ws2_Colorectal.title="Variant_calls"
        ws4_Colorectal.title="Mutations and SNPs"
        ws5_Colorectal.title="hotspots.gaps"
        ws6_Colorectal.title="Report"
        ws7_Colorectal.title="NTC variant"
        ws9_Colorectal.title="Subpanel NTC check"
        ws10_Colorectal.title="Subpanel coverage"



        gaps, ws5_output=get_gaps_file("Colorectal", path, "tester", ws5_Colorectal, wb_Colorectal, "tests")


        self.assertEqual(ws5_output["A1"].value, '1')
        self.assertEqual(ws5_output["B1"].value, "start1")
        self.assertEqual(ws5_output["C1"].value, "end1")
        self.assertEqual(ws5_output["D1"].value, "Colorectal_gap1")
        self.assertEqual(ws5_output["E1"].value, '12.0')
        self.assertEqual(ws5_output["F1"].value, '11.0')

        self.assertEqual(ws5_output["A2"].value, 2)
        self.assertEqual(ws5_output["B2"].value, "start2")
        self.assertEqual(ws5_output["C2"].value, "end2")
        self.assertEqual(ws5_output["D2"].value, "Colorectal_gap2")
        self.assertEqual(ws5_output["E2"].value, 1.0)
        self.assertEqual(ws5_output["F2"].value, 6.0)

        self.assertEqual(ws5_output["A3"].value, 3)
        self.assertEqual(ws5_output["B3"].value, "start3")
        self.assertEqual(ws5_output["C3"].value, "end3")
        self.assertEqual(ws5_output["D3"].value, "Colorectal_gap3")
        self.assertEqual(ws5_output["E3"].value, 21.0)
        self.assertEqual(ws5_output["F3"].value, 18.0)

        self.assertEqual(ws5_output["A4"].value, None)
        self.assertEqual(ws5_output["B4"].value, None)
        self.assertEqual(ws5_output["C4"].value, None)
        self.assertEqual(ws5_output["D4"].value, None)
        self.assertEqual(ws5_output["E4"].value, None)
        self.assertEqual(ws5_output["F4"].value, None)



	#Glioma
        wb_Glioma=Workbook()
        ws1_Glioma= wb_Glioma.create_sheet("Sheet_1")
        ws9_Glioma= wb_Glioma.create_sheet("Sheet_9")
        ws2_Glioma= wb_Glioma.create_sheet("Sheet_2")
        ws4_Glioma= wb_Glioma.create_sheet("Sheet_4")
        ws5_Glioma= wb_Glioma.create_sheet("Sheet_5")
        ws6_Glioma= wb_Glioma.create_sheet("Sheet_6")
        ws7_Glioma= wb_Glioma.create_sheet("Sheet_7")
        ws10_Glioma= wb_Glioma.create_sheet("Sheet_10")

        #name the tabs
        ws1_Glioma.title="Patient demographics"
        ws2_Glioma.title="Variant_calls"
        ws4_Glioma.title="Mutations and SNPs"
        ws5_Glioma.title="hotspots.gaps"
        ws6_Glioma.title="Report"
        ws7_Glioma.title="NTC variant"
        ws9_Glioma.title="Subpanel NTC check"
        ws10_Glioma.title="Subpanel coverage"


        gaps, ws5_output=get_gaps_file("Glioma", path, "tester", ws5_Glioma, wb_Glioma, "tests")


        self.assertEqual(ws5_output["A1"].value, 'No gaps')



	#Lung


        wb_Lung=Workbook()
        ws1_Lung= wb_Lung.create_sheet("Sheet_1")
        ws9_Lung= wb_Lung.create_sheet("Sheet_9")
        ws2_Lung= wb_Lung.create_sheet("Sheet_2")
        ws4_Lung= wb_Lung.create_sheet("Sheet_4")
        ws5_Lung= wb_Lung.create_sheet("Sheet_5")
        ws6_Lung= wb_Lung.create_sheet("Sheet_6")
        ws7_Lung= wb_Lung.create_sheet("Sheet_7")
        ws10_Lung= wb_Lung.create_sheet("Sheet_10")

        #name the tabs
        ws1_Lung.title="Patient demographics"
        ws2_Lung.title="Variant_calls"
        ws4_Lung.title="Mutations and SNPs"
        ws5_Lung.title="hotspots.gaps"
        ws6_Lung.title="Report"
        ws7_Lung.title="NTC variant"
        ws9_Lung.title="Subpanel NTC check"
        ws10_Lung.title="Subpanel coverage"



        gaps, ws5_output=get_gaps_file("Lung", path, "tester", ws5_Lung, wb_Lung, "tests")


        self.assertEqual(ws5_output["A1"].value, '1')
        self.assertEqual(ws5_output["B1"].value, "start1")
        self.assertEqual(ws5_output["C1"].value, "end1")
        self.assertEqual(ws5_output["D1"].value, "Lung_gap1")
        self.assertEqual(ws5_output["E1"].value, '7.0')
        self.assertEqual(ws5_output["F1"].value, '3.0')



        self.assertEqual(ws5_output["A2"].value, None)
        self.assertEqual(ws5_output["B2"].value, None)
        self.assertEqual(ws5_output["C2"].value, None)
        self.assertEqual(ws5_output["D2"].value, None)
        self.assertEqual(ws5_output["E2"].value, None)
        self.assertEqual(ws5_output["F2"].value, None)


	#Melanoma


        wb_Melanoma=Workbook()
        ws1_Melanoma= wb_Melanoma.create_sheet("Sheet_1")
        ws9_Melanoma= wb_Melanoma.create_sheet("Sheet_9")
        ws2_Melanoma= wb_Melanoma.create_sheet("Sheet_2")
        ws4_Melanoma= wb_Melanoma.create_sheet("Sheet_4")
        ws5_Melanoma= wb_Melanoma.create_sheet("Sheet_5")
        ws6_Melanoma= wb_Melanoma.create_sheet("Sheet_6")
        ws7_Melanoma= wb_Melanoma.create_sheet("Sheet_7")
        ws10_Melanoma= wb_Melanoma.create_sheet("Sheet_10")

        #name the tabs
        ws1_Melanoma.title="Patient demographics"
        ws2_Melanoma.title="Variant_calls"
        ws4_Melanoma.title="Mutations and SNPs"
        ws5_Melanoma.title="hotspots.gaps"
        ws6_Melanoma.title="Report"
        ws7_Melanoma.title="NTC variant"
        ws9_Melanoma.title="Subpanel NTC check"
        ws10_Melanoma.title="Subpanel coverage"



        gaps, ws5_output=get_gaps_file("Melanoma", path, "tester", ws5_Melanoma, wb_Melanoma, "tests")



        self.assertEqual(ws5_output["A1"].value, '3')
        self.assertEqual(ws5_output["B1"].value, "start3")
        self.assertEqual(ws5_output["C1"].value, "end3")
        self.assertEqual(ws5_output["D1"].value, "Melanoma_gap3")
        self.assertEqual(ws5_output["E1"].value, '27.0')
        self.assertEqual(ws5_output["F1"].value, '19.0')



        self.assertEqual(ws5_output["A2"].value, None)
        self.assertEqual(ws5_output["B2"].value, None)
        self.assertEqual(ws5_output["C2"].value, None)
        self.assertEqual(ws5_output["D2"].value, None)
        self.assertEqual(ws5_output["E2"].value, None)
        self.assertEqual(ws5_output["F2"].value, None)



	#Thyroid


        wb_Thyroid=Workbook()
        ws1_Thyroid= wb_Thyroid.create_sheet("Sheet_1")
        ws9_Thyroid= wb_Thyroid.create_sheet("Sheet_9")
        ws2_Thyroid= wb_Thyroid.create_sheet("Sheet_2")
        ws4_Thyroid= wb_Thyroid.create_sheet("Sheet_4")
        ws5_Thyroid= wb_Thyroid.create_sheet("Sheet_5")
        ws6_Thyroid= wb_Thyroid.create_sheet("Sheet_6")
        ws7_Thyroid= wb_Thyroid.create_sheet("Sheet_7")
        ws10_Thyroid=wb_Thyroid.create_sheet("Sheet_10")

        #name the tabs
        ws1_Thyroid.title="Patient demographics"
        ws2_Thyroid.title="Variant_calls"
        ws4_Thyroid.title="Mutations and SNPs"
        ws5_Thyroid.title="hotspots.gaps"
        ws6_Thyroid.title="Report"
        ws7_Thyroid.title="NTC variant"
        ws9_Thyroid.title="Subpanel NTC check"
        ws10_Thyroid.title="Subpanel coverage"



        gaps, ws5_output=get_gaps_file("Thyroid", path, "tester", ws5_Thyroid, wb_Thyroid, "tests")


        self.assertEqual(ws5_output["A1"].value, 'No gaps')




    def test_get_hotspots_coverage_file(self):

	#Colorectal
        Coverage= get_hotspots_coverage_file("Colorectal", path, "tester", "tests")


        self.assertEqual(Coverage.iloc[0,0], "Colorectal1")
        self.assertEqual(Coverage.iloc[0,1], 251.0)
        self.assertEqual(Coverage.iloc[0,2], 91.0)

        self.assertEqual(Coverage.iloc[1,0], "Colorectal2")
        self.assertEqual(Coverage.iloc[1,1], 252.0)
        self.assertEqual(Coverage.iloc[1,2], 92.0)

        self.assertEqual(Coverage.iloc[2,0], "Colorectal3")
        self.assertEqual(Coverage.iloc[2,1], 253.0)
        self.assertEqual(Coverage.iloc[2,2], 93.0)



	#Glioma
        Coverage= get_hotspots_coverage_file("Glioma", path, "tester", "tests")


        self.assertEqual(Coverage.iloc[0,0], "Glioma3")
        self.assertEqual(Coverage.iloc[0,1], 14.0)
        self.assertEqual(Coverage.iloc[0,2], 5.0)

        self.assertEqual(Coverage.iloc[1,0], "Glioma4")
        self.assertEqual(Coverage.iloc[1,1], 176.0)
        self.assertEqual(Coverage.iloc[1,2], 78.0)

        self.assertEqual(Coverage.iloc[2,0], "Glioma5")
        self.assertEqual(Coverage.iloc[2,1], 437.0)
        self.assertEqual(Coverage.iloc[2,2], 99.0)


        self.assertEqual(Coverage.iloc[3,0], "Glioma6")
        self.assertEqual(Coverage.iloc[3,1], 25.0)
        self.assertEqual(Coverage.iloc[3,2], 3.0)


	#Lung
        Coverage= get_hotspots_coverage_file("Lung", path, "tester", "tests")


        self.assertEqual(Coverage.iloc[0,0], "Lung8")
        self.assertEqual(Coverage.iloc[0,1], 85.0)
        self.assertEqual(Coverage.iloc[0,2], 15.0)

        self.assertEqual(Coverage.iloc[1,0], "Lung2")
        self.assertEqual(Coverage.iloc[1,1], 152.0)
        self.assertEqual(Coverage.iloc[1,2], 75.0)



	#Melanoma
        Coverage= get_hotspots_coverage_file("Melanoma", path, "tester", "tests")


        self.assertEqual(Coverage.iloc[0,0], "Melanoma9")
        self.assertEqual(Coverage.iloc[0,1], 72.0)
        self.assertEqual(Coverage.iloc[0,2], 34.0)

        self.assertEqual(Coverage.iloc[1,0], "Melanoma6")
        self.assertEqual(Coverage.iloc[1,1], 643.0)
        self.assertEqual(Coverage.iloc[1,2], 100.0)

        self.assertEqual(Coverage.iloc[2,0], "Melanoma4")
        self.assertEqual(Coverage.iloc[2,1], 27.0)
        self.assertEqual(Coverage.iloc[2,2], 12.0)




	#Thyroid
        Coverage= get_hotspots_coverage_file("Thyroid", path, "tester", "tests")


        self.assertEqual(Coverage.iloc[0,0], "Thyroid4")
        self.assertEqual(Coverage.iloc[0,1], 45.0)
        self.assertEqual(Coverage.iloc[0,2], 2.0)

        self.assertEqual(Coverage.iloc[1,0], "Thyroid9")
        self.assertEqual(Coverage.iloc[1,1], 22.0)
        self.assertEqual(Coverage.iloc[1,2], 1.0)



   
    def test_get_NTC_hotspots_coverage_file(self):

	#Colorectal

        NTC_Coverage=get_NTC_hotspots_coverage_file("Colorectal", path, "NTC", "tests")


        self.assertEqual(NTC_Coverage.iloc[0,0], 1)
        self.assertEqual(NTC_Coverage.iloc[0,1], "start1")
        self.assertEqual(NTC_Coverage.iloc[0,2], "end1")
        self.assertEqual(NTC_Coverage.iloc[0,3], "Colorectal1")
        self.assertEqual(NTC_Coverage.iloc[0,4], 77.0)
        self.assertEqual(NTC_Coverage.iloc[0,5], 54.0)

        self.assertEqual(NTC_Coverage.iloc[1,0], 2)
        self.assertEqual(NTC_Coverage.iloc[1,1], "start2")
        self.assertEqual(NTC_Coverage.iloc[1,2], "end2")
        self.assertEqual(NTC_Coverage.iloc[1,3], "Colorectal2")
        self.assertEqual(NTC_Coverage.iloc[1,4], 270.0)
        self.assertEqual(NTC_Coverage.iloc[1,5], 89.0)


	#Glioma

        NTC_Coverage=get_NTC_hotspots_coverage_file("Glioma", path, "NTC", "tests")


        self.assertEqual(NTC_Coverage.iloc[0,0], 1)
        self.assertEqual(NTC_Coverage.iloc[0,1], "start1")
        self.assertEqual(NTC_Coverage.iloc[0,2], "end1")
        self.assertEqual(NTC_Coverage.iloc[0,3], "Glioma1")
        self.assertEqual(NTC_Coverage.iloc[0,4], 76.0)
        self.assertEqual(NTC_Coverage.iloc[0,5], 34.0)

        self.assertEqual(NTC_Coverage.iloc[1,0], 2)
        self.assertEqual(NTC_Coverage.iloc[1,1], "start2")
        self.assertEqual(NTC_Coverage.iloc[1,2], "end2")
        self.assertEqual(NTC_Coverage.iloc[1,3], "Glioma2")
        self.assertEqual(NTC_Coverage.iloc[1,4], 20.0)
        self.assertEqual(NTC_Coverage.iloc[1,5], 12.0)

        self.assertEqual(NTC_Coverage.iloc[2,0], 3)
        self.assertEqual(NTC_Coverage.iloc[2,1], "start3")
        self.assertEqual(NTC_Coverage.iloc[2,2], "end3")
        self.assertEqual(NTC_Coverage.iloc[2,3], "Glioma3")
        self.assertEqual(NTC_Coverage.iloc[2,4], 79.0)
        self.assertEqual(NTC_Coverage.iloc[2,5], 36.0)

        self.assertEqual(NTC_Coverage.iloc[3,0], 4)
        self.assertEqual(NTC_Coverage.iloc[3,1], "start4")
        self.assertEqual(NTC_Coverage.iloc[3,2], "end4")
        self.assertEqual(NTC_Coverage.iloc[3,3], "Glioma4")
        self.assertEqual(NTC_Coverage.iloc[3,4], 56.0)
        self.assertEqual(NTC_Coverage.iloc[3,5], 31.0)


	#Lung

        NTC_Coverage=get_NTC_hotspots_coverage_file("Lung", path, "NTC", "tests")


        self.assertEqual(NTC_Coverage.iloc[0,0], 1)
        self.assertEqual(NTC_Coverage.iloc[0,1], "start1")
        self.assertEqual(NTC_Coverage.iloc[0,2], "end1")
        self.assertEqual(NTC_Coverage.iloc[0,3], "Lung1")
        self.assertEqual(NTC_Coverage.iloc[0,4], 7.0)
        self.assertEqual(NTC_Coverage.iloc[0,5], 3.0)

        self.assertEqual(NTC_Coverage.iloc[1,0], 2)
        self.assertEqual(NTC_Coverage.iloc[1,1], "start2")
        self.assertEqual(NTC_Coverage.iloc[1,2], "end2")
        self.assertEqual(NTC_Coverage.iloc[1,3], "Lung2")
        self.assertEqual(NTC_Coverage.iloc[1,4], 26.0)
        self.assertEqual(NTC_Coverage.iloc[1,5], 16.0)

        self.assertEqual(NTC_Coverage.iloc[2,0], 3)
        self.assertEqual(NTC_Coverage.iloc[2,1], "start3")
        self.assertEqual(NTC_Coverage.iloc[2,2], "end3")
        self.assertEqual(NTC_Coverage.iloc[2,3], "Lung3")
        self.assertEqual(NTC_Coverage.iloc[2,4], 92.0)
        self.assertEqual(NTC_Coverage.iloc[2,5], 57.0)

 


	#Melanoma

        NTC_Coverage=get_NTC_hotspots_coverage_file("Melanoma", path, "NTC", "tests")


        self.assertEqual(NTC_Coverage.iloc[0,0], 1)
        self.assertEqual(NTC_Coverage.iloc[0,1], "start1")
        self.assertEqual(NTC_Coverage.iloc[0,2], "end1")
        self.assertEqual(NTC_Coverage.iloc[0,3], "Melanoma1")
        self.assertEqual(NTC_Coverage.iloc[0,4], 54.0)
        self.assertEqual(NTC_Coverage.iloc[0,5], 34.0)

        self.assertEqual(NTC_Coverage.iloc[1,0], 2)
        self.assertEqual(NTC_Coverage.iloc[1,1], "start2")
        self.assertEqual(NTC_Coverage.iloc[1,2], "end2")
        self.assertEqual(NTC_Coverage.iloc[1,3], "Melanoma2")
        self.assertEqual(NTC_Coverage.iloc[1,4], 269.0)
        self.assertEqual(NTC_Coverage.iloc[1,5], 95.0)




	#Thyroid
	

        NTC_Coverage=get_NTC_hotspots_coverage_file("Thyroid", path, "NTC", "tests")


        self.assertEqual(NTC_Coverage.iloc[0,0], 1)
        self.assertEqual(NTC_Coverage.iloc[0,1], "start1")
        self.assertEqual(NTC_Coverage.iloc[0,2], "end1")
        self.assertEqual(NTC_Coverage.iloc[0,3], "Thyroid1")
        self.assertEqual(NTC_Coverage.iloc[0,4], 7.0)
        self.assertEqual(NTC_Coverage.iloc[0,5], 11.0)

        self.assertEqual(NTC_Coverage.iloc[1,0], 2)
        self.assertEqual(NTC_Coverage.iloc[1,1], "start2")
        self.assertEqual(NTC_Coverage.iloc[1,2], "end2")
        self.assertEqual(NTC_Coverage.iloc[1,3], "Thyroid2")
        self.assertEqual(NTC_Coverage.iloc[1,4], 27.0)
        self.assertEqual(NTC_Coverage.iloc[1,5], 15.0)

        self.assertEqual(NTC_Coverage.iloc[2,0], 3)
        self.assertEqual(NTC_Coverage.iloc[2,1], "start3")
        self.assertEqual(NTC_Coverage.iloc[2,2], "end3")
        self.assertEqual(NTC_Coverage.iloc[2,3], "Thyroid3")
        self.assertEqual(NTC_Coverage.iloc[2,4], 99.0)
        self.assertEqual(NTC_Coverage.iloc[2,5], 32.0)


        self.assertEqual(NTC_Coverage.iloc[3,0], 4)
        self.assertEqual(NTC_Coverage.iloc[3,1], "start4")
        self.assertEqual(NTC_Coverage.iloc[3,2], "end4")
        self.assertEqual(NTC_Coverage.iloc[3,3], "Thyroid4")
        self.assertEqual(NTC_Coverage.iloc[3,4], 264.0)
        self.assertEqual(NTC_Coverage.iloc[3,5], 88.0)


    def test_get_subpanel_coverage(self):



        wb_Colorectal=Workbook()
        ws1_Colorectal= wb_Colorectal.create_sheet("Sheet_1")
        ws9_Colorectal= wb_Colorectal.create_sheet("Sheet_9")
        ws2_Colorectal= wb_Colorectal.create_sheet("Sheet_2")
        ws4_Colorectal= wb_Colorectal.create_sheet("Sheet_4")
        ws5_Colorectal= wb_Colorectal.create_sheet("Sheet_5")
        ws6_Colorectal= wb_Colorectal.create_sheet("Sheet_6")
        ws7_Colorectal= wb_Colorectal.create_sheet("Sheet_7")
        ws10_Colorectal=wb_Colorectal.create_sheet("Sheet_10")

        #name the tabs
        ws1_Colorectal.title="Patient demographics"
        ws2_Colorectal.title="Variant_calls"
        ws4_Colorectal.title="Mutations and SNPs"
        ws5_Colorectal.title="hotspots.gaps"
        ws6_Colorectal.title="Report"
        ws7_Colorectal.title="NTC variant"
        ws9_Colorectal.title="Subpanel NTC check"
        ws10_Colorectal.title="Subpanel coverage"


        coverage2, ws10=get_subpanel_coverage("Colorectal", path, "tester", "tests", ws10_Colorectal)

        self.assertEqual(ws10["A2"].value, "tests____tester_Colorectal_Gene1")
        self.assertEqual(ws10["B2"].value, 50)
        self.assertEqual(ws10["C2"].value, 12)

        self.assertEqual(ws10["A3"].value, "tests____tester_Colorectal_Gene2")
        self.assertEqual(ws10["B3"].value, 20)
        self.assertEqual(ws10["C3"].value, 78)


        self.assertEqual(ws10["A4"].value, "tests____tester_Colorectal_Gene3")
        self.assertEqual(ws10["B4"].value, 100)
        self.assertEqual(ws10["C4"].value, 80)


        self.assertEqual(ws10["A5"].value, None)
        self.assertEqual(ws10["B5"].value, None)
        self.assertEqual(ws10["C5"].value, None)




	#Glioma

        wb_Glioma=Workbook()
        ws1_Glioma= wb_Glioma.create_sheet("Sheet_1")
        ws9_Glioma= wb_Glioma.create_sheet("Sheet_9")
        ws2_Glioma= wb_Glioma.create_sheet("Sheet_2")
        ws4_Glioma= wb_Glioma.create_sheet("Sheet_4")
        ws5_Glioma= wb_Glioma.create_sheet("Sheet_5")
        ws6_Glioma= wb_Glioma.create_sheet("Sheet_6")
        ws7_Glioma= wb_Glioma.create_sheet("Sheet_7")
        ws10_Glioma=wb_Glioma.create_sheet("Sheet_10")

        #name the tabs
        ws1_Glioma.title="Patient demographics"
        ws2_Glioma.title="Variant_calls"
        ws4_Glioma.title="Mutations and SNPs"
        ws5_Glioma.title="hotspots.gaps"
        ws6_Glioma.title="Report"
        ws7_Glioma.title="NTC variant"
        ws9_Glioma.title="Subpanel NTC check"
        ws10_Glioma.title="Subpanel coverage"


        coverage2, ws10=get_subpanel_coverage("Glioma", path, "tester", "tests", ws10_Glioma)

        self.assertEqual(ws10["A2"].value, "tests____tester_Glioma_Gene1")
        self.assertEqual(ws10["B2"].value, 11)
        self.assertEqual(ws10["C2"].value, 2)

        self.assertEqual(ws10["A3"].value, "tests____tester_Glioma_Gene2")
        self.assertEqual(ws10["B3"].value, 3)
        self.assertEqual(ws10["C3"].value, 4)


        self.assertEqual(ws10["A4"].value, None)
        self.assertEqual(ws10["B4"].value, None)
        self.assertEqual(ws10["C4"].value, None)



	#Lung

        wb_Lung=Workbook()
        ws1_Lung= wb_Lung.create_sheet("Sheet_1")
        ws9_Lung= wb_Lung.create_sheet("Sheet_9")
        ws2_Lung= wb_Lung.create_sheet("Sheet_2")
        ws4_Lung= wb_Lung.create_sheet("Sheet_4")
        ws5_Lung= wb_Lung.create_sheet("Sheet_5")
        ws6_Lung= wb_Lung.create_sheet("Sheet_6")
        ws7_Lung= wb_Lung.create_sheet("Sheet_7")
        ws10_Lung=wb_Lung.create_sheet("Sheet_10")

        #name the tabs
        ws1_Lung.title="Patient demographics"
        ws2_Lung.title="Variant_calls"
        ws4_Lung.title="Mutations and SNPs"
        ws5_Lung.title="hotspots.gaps"
        ws6_Lung.title="Report"
        ws7_Lung.title="NTC variant"
        ws9_Lung.title="Subpanel NTC check"
        ws10_Lung.title="Subpanel coverage"


        coverage2, ws10=get_subpanel_coverage("Lung", path, "tester", "tests", ws10_Lung)

        self.assertEqual(ws10["A2"].value, "tests____tester_Lung_Gene1")
        self.assertEqual(ws10["B2"].value, 554)
        self.assertEqual(ws10["C2"].value, 90)


        self.assertEqual(ws10["A3"].value, None)
        self.assertEqual(ws10["B3"].value, None)
        self.assertEqual(ws10["C3"].value, None)


	#Melanoma
        wb_Melanoma=Workbook()
        ws1_Melanoma= wb_Melanoma.create_sheet("Sheet_1")
        ws9_Melanoma= wb_Melanoma.create_sheet("Sheet_9")
        ws2_Melanoma= wb_Melanoma.create_sheet("Sheet_2")
        ws4_Melanoma= wb_Melanoma.create_sheet("Sheet_4")
        ws5_Melanoma= wb_Melanoma.create_sheet("Sheet_5")
        ws6_Melanoma= wb_Melanoma.create_sheet("Sheet_6")
        ws7_Melanoma= wb_Melanoma.create_sheet("Sheet_7")
        ws10_Melanoma=wb_Melanoma.create_sheet("Sheet_10")

        #name the tabs
        ws1_Melanoma.title="Patient demographics"
        ws2_Melanoma.title="Variant_calls"
        ws4_Melanoma.title="Mutations and SNPs"
        ws5_Melanoma.title="hotspots.gaps"
        ws6_Melanoma.title="Report"
        ws7_Melanoma.title="NTC variant"
        ws9_Melanoma.title="Subpanel NTC check"
        ws10_Melanoma.title="Subpanel coverage"


        coverage2, ws10=get_subpanel_coverage("Melanoma", path, "tester", "tests", ws10_Melanoma)


        self.assertEqual(ws10["A2"].value, "tests____tester_Melanoma_Gene1")
        self.assertEqual(ws10["B2"].value, 196)
        self.assertEqual(ws10["C2"].value, 23)


        self.assertEqual(ws10["A3"].value, "tests____tester_Melanoma_Gene2")
        self.assertEqual(ws10["B3"].value, 370)
        self.assertEqual(ws10["C3"].value, 76)

        self.assertEqual(ws10["A4"].value, None)
        self.assertEqual(ws10["B4"].value, None)
        self.assertEqual(ws10["C4"].value, None)


	#Thyroid

        wb_Thyroid=Workbook()
        ws1_Thyroid= wb_Thyroid.create_sheet("Sheet_1")
        ws9_Thyroid= wb_Thyroid.create_sheet("Sheet_9")
        ws2_Thyroid= wb_Thyroid.create_sheet("Sheet_2")
        ws4_Thyroid= wb_Thyroid.create_sheet("Sheet_4")
        ws5_Thyroid= wb_Thyroid.create_sheet("Sheet_5")
        ws6_Thyroid= wb_Thyroid.create_sheet("Sheet_6")
        ws7_Thyroid= wb_Thyroid.create_sheet("Sheet_7")
        ws10_Thyroid=wb_Thyroid.create_sheet("Sheet_10")

        #name the tabs
        ws1_Thyroid.title="Patient demographics"
        ws2_Thyroid.title="Variant_calls"
        ws4_Thyroid.title="Mutations and SNPs"
        ws5_Thyroid.title="hotspots.gaps"
        ws6_Thyroid.title="Report"
        ws7_Thyroid.title="NTC variant"
        ws9_Thyroid.title="Subpanel NTC check"
        ws10_Thyroid.title="Subpanel coverage"


        coverage2, ws10=get_subpanel_coverage("Thyroid", path, "tester", "tests", ws10_Thyroid)


        self.assertEqual(ws10["A2"].value, "tests____tester_Thyroid_Gene1")
        self.assertEqual(ws10["B2"].value, 14)
        self.assertEqual(ws10["C2"].value, 2)




    def test_match_polys_and_artefacts(self):


	#Colorectal

        wb_Colorectal=Workbook()
        ws1_Colorectal= wb_Colorectal.create_sheet("Sheet_1")
        ws9_Colorectal= wb_Colorectal.create_sheet("Sheet_9")
        ws2_Colorectal= wb_Colorectal.create_sheet("Sheet_2")
        ws4_Colorectal= wb_Colorectal.create_sheet("Sheet_4")
        ws5_Colorectal= wb_Colorectal.create_sheet("Sheet_5")
        ws6_Colorectal= wb_Colorectal.create_sheet("Sheet_6")
        ws7_Colorectal= wb_Colorectal.create_sheet("Sheet_7")
        ws10_Colorectal=wb_Colorectal.create_sheet("Sheet_10")

        #name the tabs
        ws1_Colorectal.title="Patient demographics"
        ws2_Colorectal.title="Variant_calls"
        ws4_Colorectal.title="Mutations and SNPs"
        ws5_Colorectal.title="hotspots.gaps"
        ws6_Colorectal.title="Report"
        ws7_Colorectal.title="NTC variant"
        ws9_Colorectal.title="Subpanel NTC check"
        ws10_Colorectal.title="Subpanel coverage"

        ws2_Colorectal['A8']=" "
        variant_report_NTC_Colorectal=get_variantReport_NTC("Colorectal", path, "NTC", "test")
        variant_report_Colorectal=get_variant_report("Colorectal", path, "tester", "test")

        variant_report_NTC, ws7=add_extra_columns_NTC_report(variant_report_NTC_Colorectal, variant_report_Colorectal, ws7_Colorectal, wb_Colorectal, path)

        variant_report_Colorectal=expand_variant_report(variant_report_Colorectal, variant_report_NTC_Colorectal)



        variant_report_4, ws2=match_polys_and_artefacts(variant_report_Colorectal, variant_report_NTC_Colorectal, artefacts_path, ws2_Colorectal)

        self.assertEqual(ws2["A10"].value, None)
        self.assertEqual(ws2["B10"].value, None)
        self.assertEqual(ws2["C10"].value, None)
        self.assertEqual(ws2["D10"].value, None)
        self.assertEqual(ws2["E10"].value, None)
        self.assertEqual(ws2["F10"].value, None)
        self.assertEqual(ws2["G10"].value, None)
        self.assertEqual(ws2["H10"].value, None)
        self.assertEqual(ws2["I10"].value, None)
        self.assertEqual(ws2["J10"].value, None)
        self.assertEqual(ws2["K10"].value, None)
        self.assertEqual(ws2["L10"].value, None)




	#Glioma

        wb_Glioma=Workbook()
        ws1_Glioma= wb_Glioma.create_sheet("Sheet_1")
        ws9_Glioma= wb_Glioma.create_sheet("Sheet_9")
        ws2_Glioma= wb_Glioma.create_sheet("Sheet_2")
        ws4_Glioma= wb_Glioma.create_sheet("Sheet_4")
        ws5_Glioma= wb_Glioma.create_sheet("Sheet_5")
        ws6_Glioma= wb_Glioma.create_sheet("Sheet_6")
        ws7_Glioma= wb_Glioma.create_sheet("Sheet_7")
        ws10_Glioma=wb_Glioma.create_sheet("Sheet_10")

        #name the tabs
        ws1_Glioma.title="Patient demographics"
        ws2_Glioma.title="Variant_calls"
        ws4_Glioma.title="Mutations and SNPs"
        ws5_Glioma.title="hotspots.gaps"
        ws6_Glioma.title="Report"
        ws7_Glioma.title="NTC variant"
        ws9_Glioma.title="Subpanel NTC check"
        ws10_Glioma.title="Subpanel coverage"

        ws2_Glioma['A8']=" "
        variant_report_NTC_Glioma=get_variantReport_NTC("Glioma", path, "NTC", "test")
        variant_report_Glioma=get_variant_report("Glioma", path, "tester", "test")

        variant_report_NTC, ws7=add_extra_columns_NTC_report(variant_report_NTC_Glioma, variant_report_Glioma, ws7_Glioma, wb_Glioma, path)

        variant_report_Glioma=expand_variant_report(variant_report_Glioma, variant_report_NTC_Glioma)



        variant_report_4, ws2=match_polys_and_artefacts(variant_report_Glioma, variant_report_NTC_Glioma, artefacts_path, ws2_Glioma)

        self.assertEqual(ws2["A10"].value, "Gene1")
        self.assertEqual(ws2["B10"].value, "exon1")
        self.assertEqual(ws2["C10"].value, "HGVSv1")
        self.assertEqual(ws2["D10"].value, "HGVSp1")
        self.assertEqual(ws2["E10"].value, 1.0)
        self.assertEqual(ws2["F10"].value, "Genotype1")
        self.assertEqual(ws2["G10"].value, 5.0)
        self.assertEqual(ws2["H10"].value, "Quality1")
        self.assertEqual(ws2["I10"].value, "Transcript1")
        self.assertEqual(ws2["J10"].value, "variant1")
        self.assertEqual(ws2["K10"].value, "")
        self.assertEqual(ws2["L10"].value, "Known Poly")



        self.assertEqual(ws2["A11"].value, "Gene2")
        self.assertEqual(ws2["B11"].value, "exon2")
        self.assertEqual(ws2["C11"].value, "HGVSv2")
        self.assertEqual(ws2["D11"].value, "HGVSp2")
        self.assertEqual(ws2["E11"].value, 2.0)
        self.assertEqual(ws2["F11"].value, "Genotype2")
        self.assertEqual(ws2["G11"].value, 6.0)
        self.assertEqual(ws2["H11"].value, "Quality2")
        self.assertEqual(ws2["I11"].value, "Transcript2")
        self.assertEqual(ws2["J11"].value, "variant2")
        self.assertEqual(ws2["K11"].value, "")
        self.assertEqual(ws2["L11"].value, "")


        self.assertEqual(ws2["A12"].value, "Gene3")
        self.assertEqual(ws2["B12"].value, "exon3")
        self.assertEqual(ws2["C12"].value, "HGVSv3")
        self.assertEqual(ws2["D12"].value, "HGVSp3")
        self.assertEqual(ws2["E12"].value, 3.0)
        self.assertEqual(ws2["F12"].value, "Genotype3")
        self.assertEqual(ws2["G12"].value, 7.0)
        self.assertEqual(ws2["H12"].value, "Quality3")
        self.assertEqual(ws2["I12"].value, "Transcript3")
        self.assertEqual(ws2["J12"].value, "variant3")
        self.assertEqual(ws2["K12"].value, "")
        self.assertEqual(ws2["L12"].value, "")


        self.assertEqual(ws2["A13"].value, None)
        self.assertEqual(ws2["B13"].value, None)
        self.assertEqual(ws2["C13"].value, None)
        self.assertEqual(ws2["D13"].value, None)
        self.assertEqual(ws2["E13"].value, None)
        self.assertEqual(ws2["F13"].value, None)
        self.assertEqual(ws2["G13"].value, None)
        self.assertEqual(ws2["H13"].value, None)
        self.assertEqual(ws2["I13"].value, None)
        self.assertEqual(ws2["J13"].value, None)
        self.assertEqual(ws2["K13"].value, None)
        self.assertEqual(ws2["L13"].value, None)




	#Lung

        wb_Lung=Workbook()
        ws1_Lung= wb_Lung.create_sheet("Sheet_1")
        ws9_Lung= wb_Lung.create_sheet("Sheet_9")
        ws2_Lung= wb_Lung.create_sheet("Sheet_2")
        ws4_Lung= wb_Lung.create_sheet("Sheet_4")
        ws5_Lung= wb_Lung.create_sheet("Sheet_5")
        ws6_Lung= wb_Lung.create_sheet("Sheet_6")
        ws7_Lung= wb_Lung.create_sheet("Sheet_7")
        ws10_Lung=wb_Lung.create_sheet("Sheet_10")

        #name the tabs
        ws1_Lung.title="Patient demographics"
        ws2_Lung.title="Variant_calls"
        ws4_Lung.title="Mutations and SNPs"
        ws5_Lung.title="hotspots.gaps"
        ws6_Lung.title="Report"
        ws7_Lung.title="NTC variant"
        ws9_Lung.title="Subpanel NTC check"
        ws10_Lung.title="Subpanel coverage"

        ws2_Lung['A8']=" "
        variant_report_NTC_Lung=get_variantReport_NTC("Lung", path, "NTC", "test")
        variant_report_Lung=get_variant_report("Lung", path, "tester", "test")

        variant_report_NTC, ws7=add_extra_columns_NTC_report(variant_report_NTC_Lung, variant_report_Lung, ws7_Lung, wb_Lung, path)

        variant_report_Lung=expand_variant_report(variant_report_Lung, variant_report_NTC_Lung)



        variant_report_4, ws2=match_polys_and_artefacts(variant_report_Lung, variant_report_NTC_Lung, artefacts_path, ws2_Lung)
 
        self.assertEqual(ws2["A10"].value, "Gene1")
        self.assertEqual(ws2["B10"].value, "exon1")
        self.assertEqual(ws2["C10"].value, "HGVSv1")
        self.assertEqual(ws2["D10"].value, "HGVSp1")
        self.assertEqual(ws2["E10"].value, 1.0)
        self.assertEqual(ws2["F10"].value, "Genotype1")
        self.assertEqual(ws2["G10"].value, 5.0)
        self.assertEqual(ws2["H10"].value, "Quality1")
        self.assertEqual(ws2["I10"].value, "Transcript1")
        self.assertEqual(ws2["J10"].value, "variant1")
        self.assertEqual(ws2["K10"].value, "")
        self.assertEqual(ws2["L10"].value, "Known Poly")



        self.assertEqual(ws2["A11"].value, "Gene2")
        self.assertEqual(ws2["B11"].value, "exon2")
        self.assertEqual(ws2["C11"].value, "HGVSv2")
        self.assertEqual(ws2["D11"].value, "HGVSp2")
        self.assertEqual(ws2["E11"].value, 2.0)
        self.assertEqual(ws2["F11"].value, "Genotype2")
        self.assertEqual(ws2["G11"].value, 6.0)
        self.assertEqual(ws2["H11"].value, "Quality2")
        self.assertEqual(ws2["I11"].value, "Transcript2")
        self.assertEqual(ws2["J11"].value, "variant2")
        self.assertEqual(ws2["K11"].value, "")
        self.assertEqual(ws2["L11"].value, "")


        self.assertEqual(ws2["A12"].value, "Gene3")
        self.assertEqual(ws2["B12"].value, "exon3")
        self.assertEqual(ws2["C12"].value, "HGVSv3")
        self.assertEqual(ws2["D12"].value, "HGVSp3")
        self.assertEqual(ws2["E12"].value, 3.0)
        self.assertEqual(ws2["F12"].value, "Genotype3")
        self.assertEqual(ws2["G12"].value, 7.0)
        self.assertEqual(ws2["H12"].value, "Quality3")
        self.assertEqual(ws2["I12"].value, "Transcript3")
        self.assertEqual(ws2["J12"].value, "variant3")
        self.assertEqual(ws2["K12"].value, "")
        self.assertEqual(ws2["L12"].value, "")



        self.assertEqual(ws2["A13"].value, "Gene4")
        self.assertEqual(ws2["B13"].value, "exon4")
        self.assertEqual(ws2["C13"].value, "HGVSv4")
        self.assertEqual(ws2["D13"].value, "HGVSp4")
        self.assertEqual(ws2["E13"].value, 4.0)
        self.assertEqual(ws2["F13"].value, "Genotype4")
        self.assertEqual(ws2["G13"].value, 8.0)
        self.assertEqual(ws2["H13"].value, "Quality4")
        self.assertEqual(ws2["I13"].value, "Transcript4")
        self.assertEqual(ws2["J13"].value, "variant4")
        self.assertEqual(ws2["K13"].value, "")
        self.assertEqual(ws2["L13"].value, "")


        self.assertEqual(ws2["A14"].value, "Gene5")
        self.assertEqual(ws2["B14"].value, "exon5")
        self.assertEqual(ws2["C14"].value, "HGVSv5")
        self.assertEqual(ws2["D14"].value, "HGVSp5")
        self.assertEqual(ws2["E14"].value, 5.0)
        self.assertEqual(ws2["F14"].value, "Genotype5")
        self.assertEqual(ws2["G14"].value, 9.0)
        self.assertEqual(ws2["H14"].value, "Quality5")
        self.assertEqual(ws2["I14"].value, "Transcript5")
        self.assertEqual(ws2["J14"].value, "variant5")
        self.assertEqual(ws2["K14"].value, "")
        self.assertEqual(ws2["L14"].value, "")


        self.assertEqual(ws2["A15"].value, None)
        self.assertEqual(ws2["B15"].value, None)
        self.assertEqual(ws2["C15"].value, None)
        self.assertEqual(ws2["D15"].value, None)
        self.assertEqual(ws2["E15"].value, None)
        self.assertEqual(ws2["F15"].value, None)
        self.assertEqual(ws2["G15"].value, None)
        self.assertEqual(ws2["H15"].value, None)
        self.assertEqual(ws2["I15"].value, None)
        self.assertEqual(ws2["J15"].value, None)
        self.assertEqual(ws2["K15"].value, None)
        self.assertEqual(ws2["L15"].value, None)




	#Thyroid

        wb_Thyroid=Workbook()
        ws1_Thyroid= wb_Thyroid.create_sheet("Sheet_1")
        ws9_Thyroid= wb_Thyroid.create_sheet("Sheet_9")
        ws2_Thyroid= wb_Thyroid.create_sheet("Sheet_2")
        ws4_Thyroid= wb_Thyroid.create_sheet("Sheet_4")
        ws5_Thyroid= wb_Thyroid.create_sheet("Sheet_5")
        ws6_Thyroid= wb_Thyroid.create_sheet("Sheet_6")
        ws7_Thyroid= wb_Thyroid.create_sheet("Sheet_7")
        ws10_Thyroid=wb_Thyroid.create_sheet("Sheet_10")

        #name the tabs
        ws1_Thyroid.title="Patient demographics"
        ws2_Thyroid.title="Variant_calls"
        ws4_Thyroid.title="Mutations and SNPs"
        ws5_Thyroid.title="hotspots.gaps"
        ws6_Thyroid.title="Report"
        ws7_Thyroid.title="NTC variant"
        ws9_Thyroid.title="Subpanel NTC check"
        ws10_Thyroid.title="Subpanel coverage"

        ws2_Thyroid['A8']=" "
        variant_report_NTC_Thyroid=get_variantReport_NTC("Thyroid", path, "NTC", "test")
        variant_report_Thyroid=get_variant_report("Thyroid", path, "tester", "test")

        variant_report_NTC, ws7=add_extra_columns_NTC_report(variant_report_NTC_Thyroid, variant_report_Thyroid, ws7_Thyroid, wb_Thyroid, path)

        variant_report_Thyroid=expand_variant_report(variant_report_Thyroid, variant_report_NTC_Thyroid)



        variant_report_4, ws2=match_polys_and_artefacts(variant_report_Thyroid, variant_report_NTC_Thyroid, artefacts_path, ws2_Thyroid)
 
        self.assertEqual(ws2["A10"].value, "Gene3")
        self.assertEqual(ws2["B10"].value, "exon3")
        self.assertEqual(ws2["C10"].value, "HGVSv3")
        self.assertEqual(ws2["D10"].value, "HGVSp3")
        self.assertEqual(ws2["E10"].value, 3.0)
        self.assertEqual(ws2["F10"].value, "Genotype3")
        self.assertEqual(ws2["G10"].value, 5.0)
        self.assertEqual(ws2["H10"].value, "Quality3")
        self.assertEqual(ws2["I10"].value, "Transcript3")
        self.assertEqual(ws2["J10"].value, "variant3")
        self.assertEqual(ws2["K10"].value, "")
        self.assertEqual(ws2["L10"].value, "")

 
        self.assertEqual(ws2["A11"].value, "Gene2")
        self.assertEqual(ws2["B11"].value, "exon2")
        self.assertEqual(ws2["C11"].value, "HGVSv2")
        self.assertEqual(ws2["D11"].value, "HGVSp2")
        self.assertEqual(ws2["E11"].value, 2.0)
        self.assertEqual(ws2["F11"].value, "Genotype2")
        self.assertEqual(ws2["G11"].value, 6.0)
        self.assertEqual(ws2["H11"].value, "Quality2")
        self.assertEqual(ws2["I11"].value, "Transcript2")
        self.assertEqual(ws2["J11"].value, "variant2")
        self.assertEqual(ws2["K11"].value, "")
        self.assertEqual(ws2["L11"].value, "")


 
        self.assertEqual(ws2["A12"].value, "Gene1")
        self.assertEqual(ws2["B12"].value, "exon1")
        self.assertEqual(ws2["C12"].value, "HGVSv1")
        self.assertEqual(ws2["D12"].value, "HGVSp1")
        self.assertEqual(ws2["E12"].value, 1.0)
        self.assertEqual(ws2["F12"].value, "Genotype1")
        self.assertEqual(ws2["G12"].value, 7.0)
        self.assertEqual(ws2["H12"].value, "Quality1")
        self.assertEqual(ws2["I12"].value, "Transcript1")
        self.assertEqual(ws2["J12"].value, "variant1")
        self.assertEqual(ws2["K12"].value, "")
        self.assertEqual(ws2["L12"].value, "Known Poly")


 
        self.assertEqual(ws2["A13"].value, "Gene7")
        self.assertEqual(ws2["B13"].value, "exon7")
        self.assertEqual(ws2["C13"].value, "HGVSv7")
        self.assertEqual(ws2["D13"].value, "HGVSp7")
        self.assertEqual(ws2["E13"].value, 7.0)
        self.assertEqual(ws2["F13"].value, "Genotype7")
        self.assertEqual(ws2["G13"].value, 11.0)
        self.assertEqual(ws2["H13"].value, "Quality7")
        self.assertEqual(ws2["I13"].value, "Transcript7")
        self.assertEqual(ws2["J13"].value, "variant7")
        self.assertEqual(ws2["K13"].value, "")
        self.assertEqual(ws2["L13"].value, "Known artefact")


        self.assertEqual(ws2["A14"].value, None)
        self.assertEqual(ws2["B14"].value, None)
        self.assertEqual(ws2["C14"].value, None)
        self.assertEqual(ws2["D14"].value, None)
        self.assertEqual(ws2["E14"].value, None)
        self.assertEqual(ws2["F14"].value, None)
        self.assertEqual(ws2["G14"].value, None)
        self.assertEqual(ws2["H14"].value, None)
        self.assertEqual(ws2["I14"].value, None)
        self.assertEqual(ws2["J14"].value, None)
        self.assertEqual(ws2["K14"].value, None)
        self.assertEqual(ws2["L14"].value, None)



