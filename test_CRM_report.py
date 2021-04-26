import unittest

from CRM_report import *

path="./tests_CRM/"

artefacts_path="./tests_CRM/"


class test_virtualhood(unittest.TestCase):


    def test_get_variantReport_NTC(self):
        self.assertEqual(len(get_variantReport_NTC("FOCUS4", path, "NTC","test")),2)
        self.assertEqual(len(get_variantReport_NTC("TP53", path, "NTC", "test")),3)

    def test_get_variant_report(self):
        self.assertEqual(len(get_variant_report("FOCUS4", path, "tester", "test")),2)
        self.assertEqual(len(get_variant_report("TP53", path, "tester", "test")),3)

    def test_add_extra_columns_NTC_report(self):

        wb_FOCUS4=Workbook()
        ws1_FOCUS4= wb_FOCUS4.create_sheet("Sheet_1")
        ws9_FOCUS4= wb_FOCUS4.create_sheet("Sheet_9")
        ws2_FOCUS4= wb_FOCUS4.create_sheet("Sheet_2")
        ws4_FOCUS4= wb_FOCUS4.create_sheet("Sheet_4")
        ws5_FOCUS4= wb_FOCUS4.create_sheet("Sheet_5")
        ws6_FOCUS4= wb_FOCUS4.create_sheet("Sheet_6")
        ws7_FOCUS4= wb_FOCUS4.create_sheet("Sheet_7")
        ws10_FOCUS4= wb_FOCUS4.create_sheet("Sheet_10")

        #name the tabs

        ws7_FOCUS4.title="NTC variant"

        variant_report_NTC_FOCUS4=get_variantReport_NTC("FOCUS4", path, "NTC", "test")
        variant_report_FOCUS4=get_variant_report("FOCUS4", path, "tester", "test")

        variant_report_NTC, ws7, wb=add_extra_columns_NTC_report(variant_report_NTC_FOCUS4, variant_report_FOCUS4, ws7_FOCUS4, wb_FOCUS4, path)
        self.assertEqual(ws7["A2"].value, "Gene1")
        self.assertEqual(ws7["B2"].value, "exon1")
        self.assertEqual(ws7["C2"].value, "HGVSv1")
        self.assertEqual(ws7["D2"].value, "HGVSp1")
        self.assertEqual(ws7["E2"].value, 2.0)
        self.assertEqual(ws7["F2"].value, "Quality1")
        self.assertEqual(ws7["G2"].value, 5)
        self.assertEqual(ws7["H2"].value, "classification")
        self.assertEqual(ws7["I2"].value, "Transcript1")
        self.assertEqual(ws7["J2"].value, "variant1")
        self.assertEqual(ws7["K2"].value, "YES")
        self.assertEqual(ws7["L2"].value, 10.0)

        self.assertEqual(ws7["A3"].value, "Gene3")
        self.assertEqual(ws7["B3"].value, "exon3")
        self.assertEqual(ws7["C3"].value, "HGVSv3")
        self.assertEqual(ws7["D3"].value, "HGVSp3")
        self.assertEqual(ws7["E3"].value, 3.0)
        self.assertEqual(ws7["F3"].value, "Quality3")
        self.assertEqual(ws7["G3"].value, 7)
        self.assertEqual(ws7["H3"].value, "classification")
        self.assertEqual(ws7["I3"].value, "Transcript3")
        self.assertEqual(ws7["J3"].value, "variant3")
        self.assertEqual(ws7["K3"].value, "NO")
        self.assertEqual(ws7["L3"].value, 21.0)

        self.assertEqual(ws7["A4"].value, None)
        self.assertEqual(ws7["B4"].value, None)
        self.assertEqual(ws7["C4"].value, None)
        self.assertEqual(ws7["D4"].value, None)
        self.assertEqual(ws7["E4"].value, None)
        self.assertEqual(ws7["F4"].value, None)
        self.assertEqual(ws7["G4"].value, None)
        self.assertEqual(ws7["H4"].value, None)
        self.assertEqual(ws7["I4"].value, None)
        self.assertEqual(ws7["J4"].value, None)
        self.assertEqual(ws7["K4"].value, None)
        self.assertEqual(ws7["L4"].value, None)

	#TP53
        wb_TP53=Workbook()
        ws1_TP53= wb_TP53.create_sheet("Sheet_1")
        ws9_TP53= wb_TP53.create_sheet("Sheet_9")
        ws2_TP53= wb_TP53.create_sheet("Sheet_2")
        ws4_TP53= wb_TP53.create_sheet("Sheet_4")
        ws5_TP53= wb_TP53.create_sheet("Sheet_5")
        ws6_TP53= wb_TP53.create_sheet("Sheet_6")
        ws7_TP53= wb_TP53.create_sheet("Sheet_7")
        ws10_TP53=wb_TP53.create_sheet("Sheet_10")

        #name the tabs

        ws7_TP53.title="NTC variant"

        variant_report_NTC_TP53=get_variantReport_NTC("TP53", path, "NTC", "test")
        variant_report_TP53=get_variant_report("TP53", path, "tester", "test")


        variant_report_NTC, ws7, wb=add_extra_columns_NTC_report(variant_report_NTC_TP53, variant_report_TP53, ws7_TP53, wb_TP53, path)
        self.assertEqual(ws7["A2"].value, "Gene1")
        self.assertEqual(ws7["B2"].value, "exon1")
        self.assertEqual(ws7["C2"].value, "HGVSv1")
        self.assertEqual(ws7["D2"].value, "HGVSp1")
        self.assertEqual(ws7["E2"].value, 1.0)
        self.assertEqual(ws7["F2"].value, "Quality1")
        self.assertEqual(ws7["G2"].value, 5)
        self.assertEqual(ws7["H2"].value, "classification")
        self.assertEqual(ws7["I2"].value, "Transcript1")
        self.assertEqual(ws7["J2"].value, "variant1")
        self.assertEqual(ws7["K2"].value, "YES")
        self.assertEqual(ws7["L2"].value, 5.0)


        self.assertEqual(ws7["A3"].value, "Gene4")
        self.assertEqual(ws7["B3"].value, "exon4")
        self.assertEqual(ws7["C3"].value, "HGVSv4")
        self.assertEqual(ws7["D3"].value, "HGVSp4")
        self.assertEqual(ws7["E3"].value, 4.0)
        self.assertEqual(ws7["F3"].value, "Quality4")
        self.assertEqual(ws7["G3"].value, 8)
        self.assertEqual(ws7["H3"].value, "classification")
        self.assertEqual(ws7["I3"].value, "Transcript4")
        self.assertEqual(ws7["J3"].value, "variant4")
        self.assertEqual(ws7["K3"].value, "NO")
        self.assertEqual(ws7["L3"].value, 32.0)


        self.assertEqual(ws7["A4"].value, "Gene5")
        self.assertEqual(ws7["B4"].value, "exon5")
        self.assertEqual(ws7["C4"].value, "HGVSv5")
        self.assertEqual(ws7["D4"].value, "HGVSp5")
        self.assertEqual(ws7["E4"].value, 5.0)
        self.assertEqual(ws7["F4"].value, "Quality5")
        self.assertEqual(ws7["G4"].value, 9)
        self.assertEqual(ws7["H4"].value, "classification")
        self.assertEqual(ws7["I4"].value, "Transcript5")
        self.assertEqual(ws7["J4"].value, "variant5")
        self.assertEqual(ws7["K4"].value, "NO")
        self.assertEqual(ws7["L4"].value, 45.0)



    def test_expand_variant_report(self):


	#FOCUS4
        variant_report_NTC_FOCUS4=get_variantReport_NTC("FOCUS4", path, "NTC", "test")
        variant_report_FOCUS4=get_variant_report("FOCUS4", path, "tester", "test")

        variant_report_FOCUS4=expand_variant_report(variant_report_FOCUS4, variant_report_NTC_FOCUS4, "FOCUS4")


        self.assertEqual(variant_report_FOCUS4.iloc[0,0], "Gene1")
        self.assertEqual(variant_report_FOCUS4.iloc[0,1], "exon1")
        self.assertEqual(variant_report_FOCUS4.iloc[0,2], "HGVSv1")
        self.assertEqual(variant_report_FOCUS4.iloc[0,3], "HGVSp1" )
        self.assertEqual(variant_report_FOCUS4.iloc[0,4], 2.0)
        self.assertEqual(variant_report_FOCUS4.iloc[0,5], "Quality1")
        self.assertEqual(variant_report_FOCUS4.iloc[0,6], 7)
        self.assertEqual(variant_report_FOCUS4.iloc[0,7], "classification")
        self.assertEqual(variant_report_FOCUS4.iloc[0,8], "Transcript1")
        self.assertEqual(variant_report_FOCUS4.iloc[0,9], "variant1")
        self.assertEqual(variant_report_FOCUS4.iloc[0,10], "")
        self.assertEqual(variant_report_FOCUS4.iloc[0,11], "")
        self.assertEqual(variant_report_FOCUS4.iloc[0,12], "")
        self.assertEqual(variant_report_FOCUS4.iloc[0,13], "")
        self.assertEqual(variant_report_FOCUS4.iloc[0,14], "")
        self.assertEqual(variant_report_FOCUS4.iloc[0,15], "")
        self.assertEqual(variant_report_FOCUS4.iloc[0,16], "")
        self.assertEqual(variant_report_FOCUS4.iloc[0,17], "14.0")
        self.assertEqual(variant_report_FOCUS4.iloc[0,18], "YES")


        self.assertEqual(variant_report_FOCUS4.iloc[1,0], "Gene7")
        self.assertEqual(variant_report_FOCUS4.iloc[1,1], "exon7")
        self.assertEqual(variant_report_FOCUS4.iloc[1,2], "HGVSv7")
        self.assertEqual(variant_report_FOCUS4.iloc[1,3], "HGVSp7" )
        self.assertEqual(variant_report_FOCUS4.iloc[1,4], 7.0)
        self.assertEqual(variant_report_FOCUS4.iloc[1,5], "Quality7")
        self.assertEqual(variant_report_FOCUS4.iloc[1,6], 11)
        self.assertEqual(variant_report_FOCUS4.iloc[1,7], "classification")
        self.assertEqual(variant_report_FOCUS4.iloc[1,8], "Transcript7")
        self.assertEqual(variant_report_FOCUS4.iloc[1,9], "variant7")
        self.assertEqual(variant_report_FOCUS4.iloc[1,10], "")
        self.assertEqual(variant_report_FOCUS4.iloc[1,11], "")
        self.assertEqual(variant_report_FOCUS4.iloc[1,12], "")
        self.assertEqual(variant_report_FOCUS4.iloc[1,13], "")
        self.assertEqual(variant_report_FOCUS4.iloc[1,14], "")
        self.assertEqual(variant_report_FOCUS4.iloc[1,15], "")
        self.assertEqual(variant_report_FOCUS4.iloc[1,16], "")
        self.assertEqual(variant_report_FOCUS4.iloc[1,17], "77.0")
        self.assertEqual(variant_report_FOCUS4.iloc[1,18], "NO")


	#TP53
        variant_report_NTC_TP53=get_variantReport_NTC("TP53", path, "NTC", "test")
        variant_report_TP53=get_variant_report("TP53", path, "tester", "test")

        variant_report_TP53=expand_variant_report(variant_report_TP53, variant_report_NTC_TP53, "TP53")


        self.assertEqual(variant_report_TP53.iloc[0,0], "Gene3")
        self.assertEqual(variant_report_TP53.iloc[0,1], "exon3")
        self.assertEqual(variant_report_TP53.iloc[0,2], "HGVSv3")
        self.assertEqual(variant_report_TP53.iloc[0,3], "HGVSp3" )
        self.assertEqual(variant_report_TP53.iloc[0,4], 3.0)
        self.assertEqual(variant_report_TP53.iloc[0,5], "Quality3")
        self.assertEqual(variant_report_TP53.iloc[0,6], 5)
        self.assertEqual(variant_report_TP53.iloc[0,7], "classification")
        self.assertEqual(variant_report_TP53.iloc[0,8], "Transcript3")
        self.assertEqual(variant_report_TP53.iloc[0,9], "1:23456A>C")
        self.assertEqual(variant_report_TP53.iloc[0,10], "1:23456")
        self.assertEqual(variant_report_TP53.iloc[0,11], "")
        self.assertEqual(variant_report_TP53.iloc[0,12], "")
        self.assertEqual(variant_report_TP53.iloc[0,13], "")
        self.assertEqual(variant_report_TP53.iloc[0,14], "")
        self.assertEqual(variant_report_TP53.iloc[0,15], "")
        self.assertEqual(variant_report_TP53.iloc[0,16], "")
        self.assertEqual(variant_report_TP53.iloc[0,17], "15.0")
        self.assertEqual(variant_report_TP53.iloc[0,18], "NO")



        self.assertEqual(variant_report_TP53.iloc[1,0], "Gene1")
        self.assertEqual(variant_report_TP53.iloc[1,1], "exon1")
        self.assertEqual(variant_report_TP53.iloc[1,2], "HGVSv1")
        self.assertEqual(variant_report_TP53.iloc[1,3], "HGVSp1" )
        self.assertEqual(variant_report_TP53.iloc[1,4], 1.0)
        self.assertEqual(variant_report_TP53.iloc[1,5], "Quality1")
        self.assertEqual(variant_report_TP53.iloc[1,6], 7)
        self.assertEqual(variant_report_TP53.iloc[1,7], "classification")
        self.assertEqual(variant_report_TP53.iloc[1,8], "Transcript1")
        self.assertEqual(variant_report_TP53.iloc[1,9], "variant1")
        self.assertEqual(variant_report_TP53.iloc[1,10], "")
        self.assertEqual(variant_report_TP53.iloc[1,11], "")
        self.assertEqual(variant_report_TP53.iloc[1,12], "")
        self.assertEqual(variant_report_TP53.iloc[1,13], "")
        self.assertEqual(variant_report_TP53.iloc[1,14], "")
        self.assertEqual(variant_report_TP53.iloc[1,15], "")
        self.assertEqual(variant_report_TP53.iloc[1,16], "")
        self.assertEqual(variant_report_TP53.iloc[1,17], "7.0")
        self.assertEqual(variant_report_TP53.iloc[1,18], "YES")

        self.assertEqual(variant_report_TP53.iloc[2,0], "Gene7")
        self.assertEqual(variant_report_TP53.iloc[2,1], "exon7")
        self.assertEqual(variant_report_TP53.iloc[2,2], "HGVSv7")
        self.assertEqual(variant_report_TP53.iloc[2,3], "HGVSp7" )
        self.assertEqual(variant_report_TP53.iloc[2,4], 7.0)
        self.assertEqual(variant_report_TP53.iloc[2,5], "Quality7")
        self.assertEqual(variant_report_TP53.iloc[2,6], 11)
        self.assertEqual(variant_report_TP53.iloc[2,7], "classification")
        self.assertEqual(variant_report_TP53.iloc[2,8], "Transcript7")
        self.assertEqual(variant_report_TP53.iloc[2,9], "variant7")
        self.assertEqual(variant_report_TP53.iloc[2,10], "")
        self.assertEqual(variant_report_TP53.iloc[2,11], "")
        self.assertEqual(variant_report_TP53.iloc[2,12], "")
        self.assertEqual(variant_report_TP53.iloc[2,13], "")
        self.assertEqual(variant_report_TP53.iloc[2,14], "")
        self.assertEqual(variant_report_TP53.iloc[2,15], "")
        self.assertEqual(variant_report_TP53.iloc[2,16], "")
        self.assertEqual(variant_report_TP53.iloc[2,17], "77.0")
        self.assertEqual(variant_report_TP53.iloc[2,18], "NO")


    def test_get_gaps_file(self):

	#FOCUS4
        wb_FOCUS4=Workbook()
        ws1_FOCUS4= wb_FOCUS4.create_sheet("Sheet_1")
        ws9_FOCUS4= wb_FOCUS4.create_sheet("Sheet_9")
        ws2_FOCUS4= wb_FOCUS4.create_sheet("Sheet_2")
        ws4_FOCUS4= wb_FOCUS4.create_sheet("Sheet_4")
        ws5_FOCUS4= wb_FOCUS4.create_sheet("Sheet_5")
        ws6_FOCUS4= wb_FOCUS4.create_sheet("Sheet_6")
        ws7_FOCUS4= wb_FOCUS4.create_sheet("Sheet_7")
        ws10_FOCUS4=wb_FOCUS4.create_sheet("Sheet_10")

        #name the tabs
        ws1_FOCUS4.title="Patient demographics"
        ws2_FOCUS4.title="Variant_calls"
        ws4_FOCUS4.title="Mutations and SNPs"
        ws5_FOCUS4.title="hotspots.gaps"
        ws6_FOCUS4.title="Report"
        ws7_FOCUS4.title="NTC variant"
        ws9_FOCUS4.title="Subpanel NTC check"
        ws10_FOCUS4.title="Subpanel coverage"

        ws5_output, ws6_output, wb=get_gaps_file("FOCUS4", path, "tester", ws5_FOCUS4, wb_FOCUS4, "tests", ws6_FOCUS4)

        self.assertEqual(ws5_output["A1"].value, '1')
        self.assertEqual(ws5_output["B1"].value, 'start1')
        self.assertEqual(ws5_output["C1"].value, 'end1')
        self.assertEqual(ws5_output["D1"].value, 'NRAS_gap1')
        self.assertEqual(ws5_output["E1"].value, '7.0')
        self.assertEqual(ws5_output["F1"].value, '3.0')

        self.assertEqual(ws5_output["A2"].value, None)
        self.assertEqual(ws5_output["B2"].value, None)
        self.assertEqual(ws5_output["C2"].value, None)
        self.assertEqual(ws5_output["D2"].value, None)
        self.assertEqual(ws5_output["E2"].value, None)
        self.assertEqual(ws5_output["F2"].value, None)


	#TP53

        wb_TP53=Workbook()
        ws1_TP53= wb_TP53.create_sheet("Sheet_1")
        ws9_TP53= wb_TP53.create_sheet("Sheet_9")
        ws2_TP53= wb_TP53.create_sheet("Sheet_2")
        ws4_TP53= wb_TP53.create_sheet("Sheet_4")
        ws5_TP53= wb_TP53.create_sheet("Sheet_5")
        ws6_TP53= wb_TP53.create_sheet("Sheet_6")
        ws7_TP53= wb_TP53.create_sheet("Sheet_7")
        ws10_TP53=wb_TP53.create_sheet("Sheet_10")

        #name the tabs
        ws1_TP53.title="Patient demographics"
        ws2_TP53.title="Variant_calls"
        ws4_TP53.title="Mutations and SNPs"
        ws5_TP53.title="hotspots.gaps"
        ws6_TP53.title="Report"
        ws7_TP53.title="NTC variant"
        ws9_TP53.title="Subpanel NTC check"
        ws10_TP53.title="Subpanel coverage"

        ws5_output, ws6_output, wb=get_gaps_file("TP53", path, "tester", ws5_TP53, wb_TP53, "tests", ws6_TP53)

        self.assertEqual(ws5_output["A1"].value, 'No gaps')


    def test_get_hotspots_coverage_file(self):

	#FOCUS4
        Coverage= get_hotspots_coverage_file("FOCUS4", path, "tester", "tests")

        self.assertEqual(Coverage.iloc[0,0], "BRAF1")
        self.assertEqual(Coverage.iloc[0,1], 251)
        self.assertEqual(Coverage.iloc[0,2], 91.0)

        self.assertEqual(Coverage.iloc[1,0], "BRAF2")
        self.assertEqual(Coverage.iloc[1,1], 252)
        self.assertEqual(Coverage.iloc[1,2], 92.0)

        self.assertEqual(Coverage.iloc[2,0], "BRAF3")
        self.assertEqual(Coverage.iloc[2,1], 253.0)
        self.assertEqual(Coverage.iloc[2,2], 93.0)

        self.assertEqual(Coverage.iloc[3,0], "KRAS2")
        self.assertEqual(Coverage.iloc[3,1], 252.0)
        self.assertEqual(Coverage.iloc[3,2], 92.0)

        self.assertEqual(Coverage.iloc[4,0], "KRAS3")
        self.assertEqual(Coverage.iloc[4,1], 253.0)
        self.assertEqual(Coverage.iloc[4,2], 93.0)

        self.assertEqual(Coverage.iloc[5,0], "NRAS3")
        self.assertEqual(Coverage.iloc[5,1], 253.0)
        self.assertEqual(Coverage.iloc[5,2], 93.0)

        self.assertEqual(Coverage.iloc[6,0], "NRAS4")
        self.assertEqual(Coverage.iloc[6,1], 254.0)
        self.assertEqual(Coverage.iloc[6,2], 94.0)

        self.assertEqual(Coverage.iloc[7,0], "NRAS5")
        self.assertEqual(Coverage.iloc[7,1], 255.0)
        self.assertEqual(Coverage.iloc[7,2], 95.0)

        self.assertEqual(Coverage.iloc[8,0], "NRAS6")
        self.assertEqual(Coverage.iloc[8,1], 256.0)
        self.assertEqual(Coverage.iloc[8,2], 96.0)

        self.assertEqual(Coverage.iloc[9,0], "PIK3CA7")
        self.assertEqual(Coverage.iloc[9,1], 257.0)
        self.assertEqual(Coverage.iloc[9,2], 97.0)

        self.assertEqual(Coverage.iloc[10,0], "PIK3CA8")
        self.assertEqual(Coverage.iloc[10,1], 258.0)
        self.assertEqual(Coverage.iloc[10,2], 98.0)

        self.assertEqual(Coverage.iloc[11,0], "PIK3CA9")
        self.assertEqual(Coverage.iloc[11,1], 259.0)
        self.assertEqual(Coverage.iloc[11,2], 99.0)

        self.assertEqual(Coverage.iloc[12,0], "PIK3CA10")
        self.assertEqual(Coverage.iloc[12,1], 260.0)
        self.assertEqual(Coverage.iloc[12,2], 100.0)

        self.assertEqual(Coverage.iloc[13,0], "TP53_1")
        self.assertEqual(Coverage.iloc[13,1], 45.0)
        self.assertEqual(Coverage.iloc[13,2], 2.0)

        self.assertEqual(Coverage.iloc[14,0], "TP53_2")
        self.assertEqual(Coverage.iloc[14,1], 22.0)
        self.assertEqual(Coverage.iloc[14,2], 1.0)


	#TP53
        Coverage= get_hotspots_coverage_file("TP53", path, "tester", "tests")

        self.assertEqual(Coverage.iloc[0,0], "TP53_1")
        self.assertEqual(Coverage.iloc[0,1], 45.0)
        self.assertEqual(Coverage.iloc[0,2], 2.0)

        self.assertEqual(Coverage.iloc[1,0], "TP53_2")
        self.assertEqual(Coverage.iloc[1,1], 22.0)
        self.assertEqual(Coverage.iloc[1,2], 1.0)


   
    def test_get_NTC_hotspots_coverage_file(self):

	#FOCUS4
        NTC_Coverage=get_NTC_hotspots_coverage_file("FOCUS4", path, "NTC", "tests")

        self.assertEqual(NTC_Coverage.iloc[0,0], 1)
        self.assertEqual(NTC_Coverage.iloc[0,1], "start1")
        self.assertEqual(NTC_Coverage.iloc[0,2], "end1")
        self.assertEqual(NTC_Coverage.iloc[0,3], "BRAF1")
        self.assertEqual(NTC_Coverage.iloc[0,4], 1.0)
        self.assertEqual(NTC_Coverage.iloc[0,5], 3.0)

        self.assertEqual(NTC_Coverage.iloc[1,0], 2)
        self.assertEqual(NTC_Coverage.iloc[1,1], "start2")
        self.assertEqual(NTC_Coverage.iloc[1,2], "end2")
        self.assertEqual(NTC_Coverage.iloc[1,3], "BRAF2")
        self.assertEqual(NTC_Coverage.iloc[1,4], 2.0)
        self.assertEqual(NTC_Coverage.iloc[1,5], 5.0)

        self.assertEqual(NTC_Coverage.iloc[2,0], 5)
        self.assertEqual(NTC_Coverage.iloc[2,1], "start3")
        self.assertEqual(NTC_Coverage.iloc[2,2], "end3")
        self.assertEqual(NTC_Coverage.iloc[2,3], "BRAF3")
        self.assertEqual(NTC_Coverage.iloc[2,4], 3.0)
        self.assertEqual(NTC_Coverage.iloc[2,5], 7.0)

        self.assertEqual(NTC_Coverage.iloc[3,0], 1)
        self.assertEqual(NTC_Coverage.iloc[3,1], "start1")
        self.assertEqual(NTC_Coverage.iloc[3,2], "end1")
        self.assertEqual(NTC_Coverage.iloc[3,3], "KRAS2")
        self.assertEqual(NTC_Coverage.iloc[3,4], 6.0)
        self.assertEqual(NTC_Coverage.iloc[3,5], 2.0)

        self.assertEqual(NTC_Coverage.iloc[4,0], 2)
        self.assertEqual(NTC_Coverage.iloc[4,1], "start2")
        self.assertEqual(NTC_Coverage.iloc[4,2], "end2")
        self.assertEqual(NTC_Coverage.iloc[4,3], "KRAS3")
        self.assertEqual(NTC_Coverage.iloc[4,4], 14.0)
        self.assertEqual(NTC_Coverage.iloc[4,5], 6.0)

        self.assertEqual(NTC_Coverage.iloc[5,0], 1)
        self.assertEqual(NTC_Coverage.iloc[5,1], "start1")
        self.assertEqual(NTC_Coverage.iloc[5,2], "end1")
        self.assertEqual(NTC_Coverage.iloc[5,3], "NRAS3")
        self.assertEqual(NTC_Coverage.iloc[5,4], 304.0)
        self.assertEqual(NTC_Coverage.iloc[5,5], 27.0)

        self.assertEqual(NTC_Coverage.iloc[6,0], 2)
        self.assertEqual(NTC_Coverage.iloc[6,1], "start2")
        self.assertEqual(NTC_Coverage.iloc[6,2], "end2")
        self.assertEqual(NTC_Coverage.iloc[6,3], "NRAS4")
        self.assertEqual(NTC_Coverage.iloc[6,4], 8.0)
        self.assertEqual(NTC_Coverage.iloc[6,5], 3.0)

        self.assertEqual(NTC_Coverage.iloc[7,0], 5)
        self.assertEqual(NTC_Coverage.iloc[7,1], "start3")
        self.assertEqual(NTC_Coverage.iloc[7,2], "end3")
        self.assertEqual(NTC_Coverage.iloc[7,3], "NRAS5")
        self.assertEqual(NTC_Coverage.iloc[7,4], 653.0)
        self.assertEqual(NTC_Coverage.iloc[7,5], 100.0)

        self.assertEqual(NTC_Coverage.iloc[8,0], 7)
        self.assertEqual(NTC_Coverage.iloc[8,1], "start4")
        self.assertEqual(NTC_Coverage.iloc[8,2], "end4")
        self.assertEqual(NTC_Coverage.iloc[8,3], "NRAS6")
        self.assertEqual(NTC_Coverage.iloc[8,4], 385.0)
        self.assertEqual(NTC_Coverage.iloc[8,5], 88.0)

        self.assertEqual(NTC_Coverage.iloc[9,0], 1)
        self.assertEqual(NTC_Coverage.iloc[9,1], "start1")
        self.assertEqual(NTC_Coverage.iloc[9,2], "end1")
        self.assertEqual(NTC_Coverage.iloc[9,3], "PIK3CA7")
        self.assertEqual(NTC_Coverage.iloc[9,4], 402.0)
        self.assertEqual(NTC_Coverage.iloc[9,5], 56.0)

        self.assertEqual(NTC_Coverage.iloc[10,0], 2)
        self.assertEqual(NTC_Coverage.iloc[10,1], "start2")
        self.assertEqual(NTC_Coverage.iloc[10,2], "end2")
        self.assertEqual(NTC_Coverage.iloc[10,3], "PIK3CA8")
        self.assertEqual(NTC_Coverage.iloc[10,4], 55.0)
        self.assertEqual(NTC_Coverage.iloc[10,5], 4.0)

        self.assertEqual(NTC_Coverage.iloc[11,0], 5)
        self.assertEqual(NTC_Coverage.iloc[11,1], "start3")
        self.assertEqual(NTC_Coverage.iloc[11,2], "end3")
        self.assertEqual(NTC_Coverage.iloc[11,3], "PIK3CA9")
        self.assertEqual(NTC_Coverage.iloc[11,4], 97.0)
        self.assertEqual(NTC_Coverage.iloc[11,5], 23.0)

        self.assertEqual(NTC_Coverage.iloc[12,0], 7)
        self.assertEqual(NTC_Coverage.iloc[12,1], "start4")
        self.assertEqual(NTC_Coverage.iloc[12,2], "end4")
        self.assertEqual(NTC_Coverage.iloc[12,3], "PIK3CA10")
        self.assertEqual(NTC_Coverage.iloc[12,4], 104.0)
        self.assertEqual(NTC_Coverage.iloc[12,5], 34.0)

        self.assertEqual(NTC_Coverage.iloc[13,0], 1)
        self.assertEqual(NTC_Coverage.iloc[13,1], "start1")
        self.assertEqual(NTC_Coverage.iloc[13,2], "end1")
        self.assertEqual(NTC_Coverage.iloc[13,3], "TP53_1")
        self.assertEqual(NTC_Coverage.iloc[13,4], 1.0)
        self.assertEqual(NTC_Coverage.iloc[13,5], 3.0)

        self.assertEqual(NTC_Coverage.iloc[14,0], 2)
        self.assertEqual(NTC_Coverage.iloc[14,1], "start2")
        self.assertEqual(NTC_Coverage.iloc[14,2], "end2")
        self.assertEqual(NTC_Coverage.iloc[14,3], "TP53_2")
        self.assertEqual(NTC_Coverage.iloc[14,4], 2.0)
        self.assertEqual(NTC_Coverage.iloc[14,5], 5.0)


	#TP53
        NTC_Coverage=get_NTC_hotspots_coverage_file("TP53", path, "NTC", "tests")


        self.assertEqual(NTC_Coverage.iloc[0,0], 1)
        self.assertEqual(NTC_Coverage.iloc[0,1], "start1")
        self.assertEqual(NTC_Coverage.iloc[0,2], "end1")
        self.assertEqual(NTC_Coverage.iloc[0,3], "TP53_1")
        self.assertEqual(NTC_Coverage.iloc[0,4], 1.0)
        self.assertEqual(NTC_Coverage.iloc[0,5], 3.0)

        self.assertEqual(NTC_Coverage.iloc[1,0], 2)
        self.assertEqual(NTC_Coverage.iloc[1,1], "start2")
        self.assertEqual(NTC_Coverage.iloc[1,2], "end2")
        self.assertEqual(NTC_Coverage.iloc[1,3], "TP53_2")
        self.assertEqual(NTC_Coverage.iloc[1,4], 2.0)
        self.assertEqual(NTC_Coverage.iloc[1,5], 5.0)




    def test_add_columns_hotspots_coverage(self):

 
        wb_FOCUS4=Workbook()
        ws1_FOCUS4= wb_FOCUS4.create_sheet("Sheet_1")
        ws9_FOCUS4= wb_FOCUS4.create_sheet("Sheet_9")
        ws2_FOCUS4= wb_FOCUS4.create_sheet("Sheet_2")
        ws4_FOCUS4= wb_FOCUS4.create_sheet("Sheet_4")
        ws5_FOCUS4= wb_FOCUS4.create_sheet("Sheet_5")
        ws6_FOCUS4= wb_FOCUS4.create_sheet("Sheet_6")
        ws7_FOCUS4= wb_FOCUS4.create_sheet("Sheet_7")
        ws10_FOCUS4= wb_FOCUS4.create_sheet("Sheet_10")


        ws1_FOCUS4.title="Patient demographics"
        ws2_FOCUS4.title="Variant_calls"
        ws4_FOCUS4.title="Mutations and SNPs"
        ws5_FOCUS4.title="hotspots.gaps"
        ws6_FOCUS4.title="Report"
        ws7_FOCUS4.title="NTC variant"
        ws9_FOCUS4.title="Subpanel NTC check"
        ws10_FOCUS4.title="Subpanel coverage"


        Coverage_FOCUS4= get_hotspots_coverage_file("FOCUS4", path, "tester", "tests")

        NTC_check=get_NTC_hotspots_coverage_file("FOCUS4", path, "NTC", "tests")

        Coverage, num_rows_coverage, ws9=add_columns_hotspots_coverage(Coverage_FOCUS4, NTC_check, ws9_FOCUS4)

        self.assertEqual(ws9["A2"].value, "BRAF1")
        self.assertEqual(ws9["B2"].value, 251)
        self.assertEqual(ws9["C2"].value, 91)
        self.assertEqual(ws9["D2"].value, 1)
        self.assertEqual(ws9["E2"].value, 0.398406374501992 )

        self.assertEqual(ws9["A3"].value, "BRAF2")
        self.assertEqual(ws9["B3"].value, 252)
        self.assertEqual(ws9["C3"].value, 92)
        self.assertEqual(ws9["D3"].value, 2)
        self.assertEqual(ws9["E3"].value, 0.7936507936507936 )

        self.assertEqual(ws9["A4"].value, "BRAF3")
        self.assertEqual(ws9["B4"].value, 253)
        self.assertEqual(ws9["C4"].value, 93)
        self.assertEqual(ws9["D4"].value, 3)
        self.assertEqual(ws9["E4"].value, 1.185770750988142 )

        self.assertEqual(ws9["A5"].value, "KRAS2")
        self.assertEqual(ws9["B5"].value, 252)
        self.assertEqual(ws9["C5"].value, 92)
        self.assertEqual(ws9["D5"].value, 6)
        self.assertEqual(ws9["E5"].value, 2.380952380952381)

        self.assertEqual(ws9["A6"].value, "KRAS3")
        self.assertEqual(ws9["B6"].value, 253)
        self.assertEqual(ws9["C6"].value, 93)
        self.assertEqual(ws9["D6"].value, 14)
        self.assertEqual(ws9["E6"].value, 5.533596837944664)

        self.assertEqual(ws9["A7"].value, "NRAS3")
        self.assertEqual(ws9["B7"].value, 253)
        self.assertEqual(ws9["C7"].value, 93)
        self.assertEqual(ws9["D7"].value, 304)
        self.assertEqual(ws9["E7"].value, 120.15810276679841 )

        self.assertEqual(ws9["A8"].value, "NRAS4")
        self.assertEqual(ws9["B8"].value, 254)
        self.assertEqual(ws9["C8"].value, 94)
        self.assertEqual(ws9["D8"].value, 8)
        self.assertEqual(ws9["E8"].value, 3.149606299212598 )

        self.assertEqual(ws9["A9"].value, "NRAS5")
        self.assertEqual(ws9["B9"].value, 255)
        self.assertEqual(ws9["C9"].value, 95)
        self.assertEqual(ws9["D9"].value, 653)
        self.assertEqual(ws9["E9"].value, 256.078431372549 )

        self.assertEqual(ws9["A10"].value, "NRAS6")
        self.assertEqual(ws9["B10"].value, 256)
        self.assertEqual(ws9["C10"].value, 96)
        self.assertEqual(ws9["D10"].value, 385)
        self.assertEqual(ws9["E10"].value, 150.390625 )

        self.assertEqual(ws9["A11"].value, "PIK3CA7")
        self.assertEqual(ws9["B11"].value, 257)
        self.assertEqual(ws9["C11"].value, 97)
        self.assertEqual(ws9["D11"].value, 402)
        self.assertEqual(ws9["E11"].value, 156.42023346303503)

        self.assertEqual(ws9["A12"].value, "PIK3CA8")
        self.assertEqual(ws9["B12"].value, 258)
        self.assertEqual(ws9["C12"].value, 98)
        self.assertEqual(ws9["D12"].value, 55)
        self.assertEqual(ws9["E12"].value, 21.31782945736434)

        self.assertEqual(ws9["A13"].value, "PIK3CA9")
        self.assertEqual(ws9["B13"].value, 259)
        self.assertEqual(ws9["C13"].value, 99)
        self.assertEqual(ws9["D13"].value, 97)
        self.assertEqual(ws9["E13"].value, 37.45173745173745 )

        self.assertEqual(ws9["A14"].value, "PIK3CA10")
        self.assertEqual(ws9["B14"].value, 260)
        self.assertEqual(ws9["C14"].value, 100)
        self.assertEqual(ws9["D14"].value, 104)
        self.assertEqual(ws9["E14"].value, 40 )

        self.assertEqual(ws9["A15"].value, "TP53_1")
        self.assertEqual(ws9["B15"].value, 45)
        self.assertEqual(ws9["C15"].value, 2.0)
        self.assertEqual(ws9["D15"].value, 1)
        self.assertEqual(ws9["E15"].value,  2.2222222222222223)

        self.assertEqual(ws9["A16"].value, "TP53_2")
        self.assertEqual(ws9["B16"].value, 22)
        self.assertEqual(ws9["C16"].value, 1)
        self.assertEqual(ws9["D16"].value, 2)
        self.assertEqual(ws9["E16"].value, 9.090909090909092 )

        self.assertEqual(ws9["A17"].value, None)
        self.assertEqual(ws9["B17"].value, None)
        self.assertEqual(ws9["C17"].value, None)
        self.assertEqual(ws9["D17"].value, None)
        self.assertEqual(ws9["E17"].value, None)


	#TP53
        wb_TP53=Workbook()
        ws1_TP53= wb_TP53.create_sheet("Sheet_1")
        ws9_TP53= wb_TP53.create_sheet("Sheet_9")
        ws2_TP53= wb_TP53.create_sheet("Sheet_2")
        ws4_TP53= wb_TP53.create_sheet("Sheet_4")
        ws5_TP53= wb_TP53.create_sheet("Sheet_5")
        ws6_TP53= wb_TP53.create_sheet("Sheet_6")
        ws7_TP53= wb_TP53.create_sheet("Sheet_7")
        ws10_TP53= wb_TP53.create_sheet("Sheet_10")


        ws1_TP53.title="Patient demographics"
        ws2_TP53.title="Variant_calls"
        ws4_TP53.title="Mutations and SNPs"
        ws5_TP53.title="hotspots.gaps"
        ws6_TP53.title="Report"
        ws7_TP53.title="NTC variant"
        ws9_TP53.title="Subpanel NTC check"
        ws10_TP53.title="Subpanel coverage"


        Coverage_TP53= get_hotspots_coverage_file("TP53", path, "tester", "tests")

        NTC_check=get_NTC_hotspots_coverage_file("TP53", path, "NTC", "tests")

        Coverage, num_rows_coverage, ws9=add_columns_hotspots_coverage(Coverage_TP53, NTC_check, ws9_TP53)

        self.assertEqual(ws9["A2"].value, "TP53_1")
        self.assertEqual(ws9["B2"].value, 45)
        self.assertEqual(ws9["C2"].value, 2.0)
        self.assertEqual(ws9["D2"].value, 1)
        self.assertEqual(ws9["E2"].value,  2.2222222222222223)

        self.assertEqual(ws9["A3"].value, "TP53_2")
        self.assertEqual(ws9["B3"].value, 22)
        self.assertEqual(ws9["C3"].value, 1)
        self.assertEqual(ws9["D3"].value, 2)
        self.assertEqual(ws9["E3"].value, 9.090909090909092 )

        self.assertEqual(ws9["A4"].value, None)
        self.assertEqual(ws9["B4"].value, None)
        self.assertEqual(ws9["C4"].value, None)
        self.assertEqual(ws9["D4"].value, None)
        self.assertEqual(ws9["E4"].value, None)


    def test_match_polys_and_artefacts(self):

	#FOCUS4

        wb_FOCUS4=Workbook()
        ws1_FOCUS4= wb_FOCUS4.create_sheet("Sheet_1")
        ws9_FOCUS4= wb_FOCUS4.create_sheet("Sheet_9")
        ws2_FOCUS4= wb_FOCUS4.create_sheet("Sheet_2")
        ws4_FOCUS4= wb_FOCUS4.create_sheet("Sheet_4")
        ws5_FOCUS4= wb_FOCUS4.create_sheet("Sheet_5")
        ws6_FOCUS4= wb_FOCUS4.create_sheet("Sheet_6")
        ws7_FOCUS4= wb_FOCUS4.create_sheet("Sheet_7")
        ws10_FOCUS4=wb_FOCUS4.create_sheet("Sheet_10")

        #name the tabs
        ws1_FOCUS4.title="Patient demographics"
        ws2_FOCUS4.title="Variant_calls"
        ws4_FOCUS4.title="Mutations and SNPs"
        ws5_FOCUS4.title="hotspots.gaps"
        ws6_FOCUS4.title="Report"
        ws7_FOCUS4.title="NTC variant"
        ws9_FOCUS4.title="Subpanel NTC check"
        ws10_FOCUS4.title="Subpanel coverage"

        ws2_FOCUS4['A8']=" "
        variant_report_NTC_FOCUS4=get_variantReport_NTC("FOCUS4", path, "NTC", "test")
        variant_report_FOCUS4=get_variant_report("FOCUS4", path, "tester", "test")

        variant_report_NTC, ws7, wb=add_extra_columns_NTC_report(variant_report_NTC_FOCUS4, variant_report_FOCUS4, ws7_FOCUS4, wb_FOCUS4, path)

        variant_report_FOCUS4=expand_variant_report(variant_report_FOCUS4, variant_report_NTC_FOCUS4, "FOCUS4")

        variant_report_4, ws2=match_polys_and_artefacts(variant_report_FOCUS4, variant_report_NTC_FOCUS4, artefacts_path, ws2_FOCUS4, "FOCUS4")

        self.assertEqual(ws2["A10"].value, "Gene1")
        self.assertEqual(ws2["B10"].value, "exon1")
        self.assertEqual(ws2["C10"].value, "HGVSv1")
        self.assertEqual(ws2["D10"].value, "HGVSp1")
        self.assertEqual(ws2["E10"].value, 2.0)
        self.assertEqual(ws2["F10"].value, "Quality1")
        self.assertEqual(ws2["G10"].value, 7.0)
        self.assertEqual(ws2["H10"].value, "classification")
        self.assertEqual(ws2["I10"].value, "Transcript1")
        self.assertEqual(ws2["J10"].value, "variant1")
        self.assertEqual(ws2["K10"].value, '')
        self.assertEqual(ws2["L10"].value, "Known Poly")
        self.assertEqual(ws2["M10"].value, 1)
        self.assertEqual(ws2["N10"].value, "Known Poly")
        self.assertEqual(ws2["O10"].value, 1)
        self.assertEqual(ws2["P10"].value, '')
        self.assertEqual(ws2["Q10"].value, '')
        self.assertEqual(ws2["R10"].value, '14.0')
        self.assertEqual(ws2["S10"].value, "YES")
        self.assertEqual(ws2["T10"].value, '14.0')
        self.assertEqual(ws2["U10"].value, 10.0)
        self.assertEqual(ws2["V10"].value, 0.7142857142857143)
        self.assertEqual(ws2["W10"].value, None)

        self.assertEqual(ws2["A11"].value, "Gene7")
        self.assertEqual(ws2["B11"].value, "exon7")
        self.assertEqual(ws2["C11"].value, "HGVSv7")
        self.assertEqual(ws2["D11"].value, "HGVSp7")
        self.assertEqual(ws2["E11"].value, 7.0)
        self.assertEqual(ws2["F11"].value, "Quality7")
        self.assertEqual(ws2["G11"].value, 11.0)
        self.assertEqual(ws2["H11"].value, "classification")
        self.assertEqual(ws2["I11"].value, "Transcript7")
        self.assertEqual(ws2["J11"].value, "variant7")
        self.assertEqual(ws2["K11"].value, '')
        self.assertEqual(ws2["L11"].value, "Known artefact")
        self.assertEqual(ws2["M11"].value, 3)
        self.assertEqual(ws2["N11"].value, "Known artefact")
        self.assertEqual(ws2["O11"].value, 3)
        self.assertEqual(ws2["P11"].value, '')
        self.assertEqual(ws2["Q11"].value, 'classification1')
        self.assertEqual(ws2["R11"].value, '77.0')
        self.assertEqual(ws2["S11"].value, 'NO')
        self.assertEqual(ws2["T11"].value, '' )
        self.assertEqual(ws2["U11"].value, '')
        self.assertEqual(ws2["V11"].value, '' )
        self.assertEqual(ws2["W11"].value, None)

        self.assertEqual(ws2["A12"].value, None)
        self.assertEqual(ws2["B12"].value, None)
        self.assertEqual(ws2["C12"].value, None)
        self.assertEqual(ws2["D12"].value, None)
        self.assertEqual(ws2["E12"].value, None)
        self.assertEqual(ws2["F12"].value, None)
        self.assertEqual(ws2["G12"].value, None)
        self.assertEqual(ws2["H12"].value, None)
        self.assertEqual(ws2["I12"].value, None)
        self.assertEqual(ws2["J12"].value, None)
        self.assertEqual(ws2["K12"].value, None)
        self.assertEqual(ws2["L12"].value, None)
        self.assertEqual(ws2["M12"].value, None)
        self.assertEqual(ws2["N12"].value, None)
        self.assertEqual(ws2["O12"].value, None)
        self.assertEqual(ws2["P12"].value, None)
        self.assertEqual(ws2["Q12"].value, None)
        self.assertEqual(ws2["R12"].value, None)
        self.assertEqual(ws2["S12"].value, None)
        self.assertEqual(ws2["T12"].value, None)
        self.assertEqual(ws2["U12"].value, None)
        self.assertEqual(ws2["V12"].value, None)
        self.assertEqual(ws2["W12"].value, None)


	#TP53
        wb_TP53=Workbook()
        ws1_TP53= wb_TP53.create_sheet("Sheet_1")
        ws9_TP53= wb_TP53.create_sheet("Sheet_9")
        ws2_TP53= wb_TP53.create_sheet("Sheet_2")
        ws4_TP53= wb_TP53.create_sheet("Sheet_4")
        ws5_TP53= wb_TP53.create_sheet("Sheet_5")
        ws6_TP53= wb_TP53.create_sheet("Sheet_6")
        ws7_TP53= wb_TP53.create_sheet("Sheet_7")
        ws10_TP53=wb_TP53.create_sheet("Sheet_10")

        #name the tabs
        ws1_TP53.title="Patient demographics"
        ws2_TP53.title="Variant_calls"
        ws4_TP53.title="Mutations and SNPs"
        ws5_TP53.title="hotspots.gaps"
        ws6_TP53.title="Report"
        ws7_TP53.title="NTC variant"
        ws9_TP53.title="Subpanel NTC check"
        ws10_TP53.title="Subpanel coverage"

        ws2_TP53['A8']=" "
        variant_report_NTC_TP53=get_variantReport_NTC("TP53", path, "NTC", "test")
        variant_report_TP53=get_variant_report("TP53", path, "tester", "test")

        variant_report_NTC, ws7, wb=add_extra_columns_NTC_report(variant_report_NTC_TP53, variant_report_TP53, ws7_TP53, wb_TP53, path)

        variant_report_TP53=expand_variant_report(variant_report_TP53, variant_report_NTC_TP53, "TP53")

        variant_report_4, ws2=match_polys_and_artefacts(variant_report_TP53, variant_report_NTC_TP53, artefacts_path, ws2_TP53, "TP53")

        self.assertEqual(ws2["A10"].value, "Gene3")
        self.assertEqual(ws2["B10"].value, "exon3")
        self.assertEqual(ws2["C10"].value, "HGVSv3")
        self.assertEqual(ws2["D10"].value, "HGVSp3")
        self.assertEqual(ws2["E10"].value, 3.0)
        self.assertEqual(ws2["F10"].value, "Quality3")
        self.assertEqual(ws2["G10"].value, 5)
        self.assertEqual(ws2["H10"].value, "classification")
        self.assertEqual(ws2["I10"].value, "Transcript3")
        self.assertEqual(ws2["J10"].value, "1:23456A>C")
        self.assertEqual(ws2["K10"].value, '1:23456')
        self.assertEqual(ws2["L10"].value, '')
        self.assertEqual(ws2["M10"].value, '')
        self.assertEqual(ws2["N10"].value, '')
        self.assertEqual(ws2["O10"].value, '')
        self.assertEqual(ws2["P10"].value, '')
        self.assertEqual(ws2["Q10"].value, '')
        self.assertEqual(ws2["R10"].value, '15.0')
        self.assertEqual(ws2["S10"].value, "NO")
        self.assertEqual(ws2["T10"].value, '' )
        self.assertEqual(ws2["U10"].value, '')
        self.assertEqual(ws2["V10"].value, '' )
        self.assertEqual(ws2["W10"].value, None)

        self.assertEqual(ws2["A11"].value, "Gene1")
        self.assertEqual(ws2["B11"].value, "exon1")
        self.assertEqual(ws2["C11"].value, "HGVSv1")
        self.assertEqual(ws2["D11"].value, "HGVSp1")
        self.assertEqual(ws2["E11"].value, 1.0)
        self.assertEqual(ws2["F11"].value, "Quality1")
        self.assertEqual(ws2["G11"].value, 7)
        self.assertEqual(ws2["H11"].value, "classification")
        self.assertEqual(ws2["I11"].value, "Transcript1")
        self.assertEqual(ws2["J11"].value, "variant1")
        self.assertEqual(ws2["K11"].value, '')
        self.assertEqual(ws2["L11"].value, 'Known Poly')
        self.assertEqual(ws2["M11"].value, 1)
        self.assertEqual(ws2["N11"].value, 'Known Poly')
        self.assertEqual(ws2["O11"].value, 1)
        self.assertEqual(ws2["P11"].value, '')
        self.assertEqual(ws2["Q11"].value, '')
        self.assertEqual(ws2["R11"].value, '7.0')
        self.assertEqual(ws2["S11"].value, "YES")
        self.assertEqual(ws2["T11"].value, '7.0' )
        self.assertEqual(ws2["U11"].value, 5.0)
        self.assertEqual(ws2["V11"].value, 0.7142857142857143)
        self.assertEqual(ws2["W11"].value, None)

        self.assertEqual(ws2["A12"].value, "Gene7")
        self.assertEqual(ws2["B12"].value, "exon7")
        self.assertEqual(ws2["C12"].value, "HGVSv7")
        self.assertEqual(ws2["D12"].value, "HGVSp7")
        self.assertEqual(ws2["E12"].value, 7.0)
        self.assertEqual(ws2["F12"].value, "Quality7")
        self.assertEqual(ws2["G12"].value, 11)
        self.assertEqual(ws2["H12"].value, "classification")
        self.assertEqual(ws2["I12"].value, "Transcript7")
        self.assertEqual(ws2["J12"].value, "variant7")
        self.assertEqual(ws2["K12"].value, '')
        self.assertEqual(ws2["L12"].value, 'Known artefact')
        self.assertEqual(ws2["M12"].value, 3)
        self.assertEqual(ws2["N12"].value, 'Known artefact')
        self.assertEqual(ws2["O12"].value, 3)
        self.assertEqual(ws2["P12"].value, '')
        self.assertEqual(ws2["Q12"].value, 'classification1')
        self.assertEqual(ws2["R12"].value, '77.0')
        self.assertEqual(ws2["S12"].value, "NO")
        self.assertEqual(ws2["T12"].value, '' )
        self.assertEqual(ws2["U12"].value, '')
        self.assertEqual(ws2["V12"].value, '' )
        self.assertEqual(ws2["W12"].value, None)


























