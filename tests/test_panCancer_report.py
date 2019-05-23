import unittest
from virtualhood import *

path ="/home/transfer/pipelines/PanCancerWorksheetCreator/tests/"

class test_virtualhood(unittest.TestCase):


    def test_get_variantReport_NTC(self):
        self.assertEqual(len(get_variantReport_NTC("Breast", path)),2)
        self.assertEqual(len(get_variantReport_NTC("Colorectal", path)),1)
        self.assertEqual(len(get_variantReport_NTC("GIST", path)),8)
        self.assertEqual(len(get_variantReport_NTC("Glioma", path)),7)
        self.assertEqual(len(get_variantReport_NTC("HN", path)),6)
        self.assertEqual(len(get_variantReport_NTC("Lung", path)),5)
        self.assertEqual(len(get_variantReport_NTC("Melanoma", path)),4)
        self.assertEqual(len(get_variantReport_NTC("Ovarian", path)),3)
        self.assertEqual(len(get_variantReport_NTC("Prostate", path)),2)
        self.assertEqual(len(get_variantReport_NTC("Thyroid", path)),2)

    def test_get_variant_report(self):
        self.assertEqual(len(get_variant_report("Breast", path, "tester")),2)
        self.assertEqual(len(get_variant_report("Colorectal", path, "tester")),0)
        self.assertEqual(len(get_variant_report("GIST", path, "tester")),0)
        self.assertEqual(len(get_variant_report("Glioma", path, "tester")),3)
        self.assertEqual(len(get_variant_report("HN", path, "tester")),4)
        self.assertEqual(len(get_variant_report("Lung", path, "tester")),6)
        self.assertEqual(len(get_variant_report("Melanoma", path, "tester")),3)
        self.assertEqual(len(get_variant_report("Ovarian", path, "tester")),5)
        self.assertEqual(len(get_variant_report("Prostate", path, "tester")),1)
        self.assertEqual(len(get_variant_report("Thyroid", path, "tester")),7)

    def test_add_extra_columns_NTC_report(self):
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Breast", path) ,get_variant_report("Breast", path, "tester")).iloc[1,10],"NO")
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Breast", path) ,get_variant_report("Breast", path, "tester")).iloc[0,11],9.00)
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Colorectal", path) ,get_variant_report("Colorectal", path, "tester")).iloc[0,10],"NO")
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Colorectal", path), get_variant_report("Colorectal", path, "tester")).iloc[0,11],5.00)
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("GIST", path) ,get_variant_report("GIST", path, "tester")).iloc[3,10],"NO")
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("GIST", path), get_variant_report("GIST", path, "tester")).iloc[4,11],325.7280)
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Glioma", path) ,get_variant_report("Glioma", path, "tester")).iloc[3,10],"NO")
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Glioma", path), get_variant_report("Glioma", path, "tester")).iloc[4,11],325.7280)        
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("HN", path) ,get_variant_report("HN", path, "tester")).iloc[5,10],"NO")
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("HN", path), get_variant_report("HN", path, "tester")).iloc[1,11],331.11839999999995)
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Lung", path) ,get_variant_report("Lung", path, "tester")).iloc[3,10],"YES")
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Lung", path), get_variant_report("Lung", path, "tester")).iloc[4,11],325.7280)  
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Melanoma", path) ,get_variant_report("Melanoma", path, "tester")).iloc[3,10],"NO")
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Melanoma", path), get_variant_report("Melanoma", path, "tester")).iloc[3,11],239.3388)
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Prostate", path) ,get_variant_report("Prostate", path, "tester")).iloc[0,10],"NO")
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Prostate", path), get_variant_report("Prostate", path, "tester")).iloc[1,11],331.11839999999995)
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Thyroid", path) ,get_variant_report("Thyroid", path, "tester")).iloc[0,10],"NO")
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Thyroid", path), get_variant_report("Thyroid", path, "tester")).iloc[0,11],14.0)
   
    def test_expand_variant_report(self):
        self.assertEqual(expand_variant_report(get_variant_report("Breast", path, "tester"), get_variantReport_NTC("Breast", path)).iloc[0,15],0)
        self.assertEqual(expand_variant_report(get_variant_report("Glioma", path, "tester"), get_variantReport_NTC("Glioma", path)).iloc[0,15],'51.953399999999995')
        self.assertEqual(expand_variant_report(get_variant_report("HN", path, "tester"), get_variantReport_NTC("HN", path)).iloc[0,15],'51.953399999999995')
        self.assertEqual(expand_variant_report(get_variant_report("Lung", path, "tester"), get_variantReport_NTC("Lung", path)).iloc[2,15],'195.8125')    
        self.assertEqual(expand_variant_report(get_variant_report("Melanoma", path, "tester"), get_variantReport_NTC("Melanoma", path)).iloc[0,15],'207.3024')
        self.assertEqual(expand_variant_report(get_variant_report("Ovarian", path, "tester"), get_variantReport_NTC("Ovarian", path)).iloc[0,15],'239.3388')
        self.assertEqual(expand_variant_report(get_variant_report("Prostate", path, "tester"), get_variantReport_NTC("Prostate", path)).iloc[0,15],'239.3388') 

    def test_get_gaps_file(self):
        self.assertEqual((len(get_gaps_file("Breast", path, "tester"))+1),14)
        self.assertEqual((len(get_gaps_file("Colorectal", path, "tester"))+1),13)
        self.assertEqual((len(get_gaps_file("GIST", path, "tester"))+1),12)
        self.assertEqual((len(get_gaps_file("Glioma", path, "tester"))+1),11)
        self.assertEqual((len(get_gaps_file("HN", path, "tester"))+1),10)
        self.assertEqual((len(get_gaps_file("Lung", path, "tester"))+1),9)
        self.assertEqual((len(get_gaps_file("Melanoma", path, "tester"))+1),8)
        self.assertEqual((len(get_gaps_file("Ovarian", path, "tester"))+1),7)
        self.assertEqual((len(get_gaps_file("Prostate", path, "tester"))+1),6)
        self.assertEqual((len(get_gaps_file("Thyroid", path, "tester"))+1),5)        

    def test_get_CNV_file(self):
        self.assertEqual(len(get_CNV_file("Breast", path, "tester")),5)
        self.assertEqual(len(get_CNV_file("Colorectal", path, "tester")),3)
        self.assertEqual(len(get_CNV_file("GIST", path, "tester")),1)
        self.assertEqual(len(get_CNV_file("Glioma", path, "tester")),2)
        self.assertEqual(len(get_CNV_file("HN", path, "tester")),4)
        self.assertEqual(len(get_CNV_file("Lung", path, "tester")),6)
        self.assertEqual(len(get_CNV_file("Melanoma", path, "tester")),0)
        self.assertEqual(len(get_CNV_file("Ovarian", path, "tester")),7)
        self.assertEqual(len(get_CNV_file("Prostate", path, "tester")),3)
        self.assertEqual(len(get_CNV_file("Thyroid", path, "tester")),1)


    def test_get_hotspots_coverage_file(self):
        self.assertEqual(len(get_hotspots_coverage_file("Breast", path, "tester")),8)
        self.assertEqual(len(get_hotspots_coverage_file("Colorectal", path, "tester")),7)
        self.assertEqual(len(get_hotspots_coverage_file("GIST", path, "tester")),6)
        self.assertEqual(len(get_hotspots_coverage_file("Glioma", path, "tester")),5)
        self.assertEqual(len(get_hotspots_coverage_file("HN", path, "tester")),4)
        self.assertEqual(len(get_hotspots_coverage_file("Lung", path, "tester")),3)
        self.assertEqual(len(get_hotspots_coverage_file("Melanoma", path, "tester")),2)
        self.assertEqual(len(get_hotspots_coverage_file("Ovarian", path, "tester")),1)
        self.assertEqual(len(get_hotspots_coverage_file("Prostate", path, "tester")),0)
        self.assertEqual(len(get_hotspots_coverage_file("Thyroid", path, "tester")),4)


    def test_get_NTC_hotspots_coverage_file(self):
        self.assertEqual(len(get_NTC_hotspots_coverage_file("Breast", path)),8)
        self.assertEqual(len(get_NTC_hotspots_coverage_file("Colorectal", path)),7)
        self.assertEqual(len(get_NTC_hotspots_coverage_file("GIST", path)),6)
        self.assertEqual(len(get_NTC_hotspots_coverage_file("Glioma", path)),5)
        self.assertEqual(len(get_NTC_hotspots_coverage_file("HN", path)),4)
        self.assertEqual(len(get_NTC_hotspots_coverage_file("Lung", path)),3)
        self.assertEqual(len(get_NTC_hotspots_coverage_file("Melanoma", path)),2)
        self.assertEqual(len(get_NTC_hotspots_coverage_file("Ovarian", path)),1)
        self.assertEqual(len(get_NTC_hotspots_coverage_file("Prostate", path)),0)
        self.assertEqual(len(get_NTC_hotspots_coverage_file("Thyroid", path)),4)

   
    def test_add_columns_hotspots_coverage(self):
        self.assertEqual(add_columns_hotspots_coverage(get_hotspots_coverage_file("Breast", path, "tester"), get_NTC_hotspots_coverage_file("Breast", path))[0].iloc[2,4], 5.818181818181818)
        self.assertEqual(add_columns_hotspots_coverage(get_hotspots_coverage_file("Colorectal", path, "tester"), get_NTC_hotspots_coverage_file("Colorectal", path))[0].iloc[2,4], 34.18181818181818)
        self.assertEqual(add_columns_hotspots_coverage(get_hotspots_coverage_file("GIST", path, "tester"), get_NTC_hotspots_coverage_file("GIST", path))[0].iloc[2,4], 25.818181818181817)
        self.assertEqual(add_columns_hotspots_coverage(get_hotspots_coverage_file("Glioma", path, "tester"), get_NTC_hotspots_coverage_file("Glioma", path))[0].iloc[1,4], 31.026993484331367)
        self.assertEqual(add_columns_hotspots_coverage(get_hotspots_coverage_file("HN", path, "tester"), get_NTC_hotspots_coverage_file("HN", path))[0].iloc[1,4], 6.211180124223603)
        self.assertEqual(add_columns_hotspots_coverage(get_hotspots_coverage_file("Lung", path, "tester"), get_NTC_hotspots_coverage_file("Lung", path))[0].iloc[1,4], 2.1739130434782608)
        self.assertEqual(add_columns_hotspots_coverage(get_hotspots_coverage_file("Melanoma", path, "tester"), get_NTC_hotspots_coverage_file("Melanoma", path))[0].iloc[1,4],2.1739130434782608)
        self.assertEqual(add_columns_hotspots_coverage(get_hotspots_coverage_file("Ovarian", path, "tester"), get_NTC_hotspots_coverage_file("Ovarian", path))[0].iloc[0,4],2.0202020202020203)
        self.assertEqual(add_columns_hotspots_coverage(get_hotspots_coverage_file("Thyroid", path, "tester"), get_NTC_hotspots_coverage_file("Thyroid", path))[0].iloc[3,4],5.105105105105105)

    def test_get_genescreen_coverage_file(self):
        self.assertEqual(len(get_genescreen_coverage_file("Breast", path, "tester")),10)
        self.assertEqual(len(get_genescreen_coverage_file("Colorectal", path, "tester")),9)
        self.assertEqual(len(get_genescreen_coverage_file("GIST", path, "tester")),8)
        self.assertEqual(len(get_genescreen_coverage_file("Glioma", path, "tester")),7)
        self.assertEqual(len(get_genescreen_coverage_file("HN", path, "tester")),6)
        self.assertEqual(len(get_genescreen_coverage_file("Lung", path, "tester")),5)
        self.assertEqual(len(get_genescreen_coverage_file("Melanoma", path, "tester")),4)
        self.assertEqual(len(get_genescreen_coverage_file("Ovarian", path, "tester")),3)
        self.assertEqual(len(get_genescreen_coverage_file("Prostate", path, "tester")),2)
        self.assertEqual(len(get_genescreen_coverage_file("Thyroid", path, "tester")),1)

    def get_NTC_genescreen_coverage_file(self):
        self.assertEqual(len(get_NTC_genescreen_coverage_file("Breast", path)),10)
        self.assertEqual(len(get_NTC_genescreen_coverage_file("Colorectal", path)),9)
        self.assertEqual(len(get_NTC_genescreen_coverage_file("GIST", path)),3)
        self.assertEqual(len(get_NTC_genescreen_coverage_file("Glioma", path)),4)
        self.assertEqual(len(get_NTC_genescreen_coverage_file("HN", path)),5)
        self.assertEqual(len(get_NTC_genescreen_coverage_file("Lung", path)),6)
        self.assertEqual(len(get_NTC_genescreen_coverage_file("Melanoma", path)),7)
        self.assertEqual(len(get_NTC_genescreen_coverage_file("Ovarian", path)),8)
        self.assertEqual(len(get_NTC_genescreen_coverage_file("Prostate", path)),9)
        self.assertEqual(len(get_NTC_genescreen_coverage_file("Thyroid", path)),10)

    def test_add_columns_genescreen_coverage_file(self):
        self.assertEqual(add_columns_genescreen_coverage(get_genescreen_coverage_file("Breast", path, "tester"), get_NTC_genescreen_coverage_file("Breast", path),8).iloc[2,4], 6.181818181818182)
        self.assertEqual(add_columns_genescreen_coverage(get_genescreen_coverage_file("Colorectal", path, "tester"), get_NTC_genescreen_coverage_file("Colorectal", path),7).iloc[7,4], 34.63203463203463)
        self.assertEqual(add_columns_genescreen_coverage(get_genescreen_coverage_file("GIST", path, "tester"), get_NTC_genescreen_coverage_file("GIST", path),6).iloc[4,4], 7.109004739336493)        
        self.assertEqual(add_columns_genescreen_coverage(get_genescreen_coverage_file("Glioma", path, "tester"), get_NTC_genescreen_coverage_file("Glioma", path),5).iloc[4,4], 0.47393364928909953)
        self.assertEqual(add_columns_genescreen_coverage(get_genescreen_coverage_file("HN", path, "tester"), get_NTC_genescreen_coverage_file("HN", path),4).iloc[4,4], 1.4218009478672986) 
        self.assertEqual(add_columns_genescreen_coverage(get_genescreen_coverage_file("Lung", path, "tester"), get_NTC_genescreen_coverage_file("Lung", path),3).iloc[4,4], 3.7914691943127963)        
        self.assertEqual(add_columns_genescreen_coverage(get_genescreen_coverage_file("Melanoma", path, "tester"), get_NTC_genescreen_coverage_file("Melanoma", path),2).iloc[3,4], 6.0402684563758395)
        self.assertEqual(add_columns_genescreen_coverage(get_genescreen_coverage_file("Ovarian", path, "tester"), get_NTC_genescreen_coverage_file("Ovarian", path),1).iloc[2,4], 0.36363636363636365)        
        self.assertEqual(add_columns_genescreen_coverage(get_genescreen_coverage_file("Prostate", path, "tester"), get_NTC_genescreen_coverage_file("Prostate", path),0).iloc[1,4], 16.149068322981368)        
        self.assertEqual(add_columns_genescreen_coverage(get_genescreen_coverage_file("Thyroid", path, "tester"), get_NTC_genescreen_coverage_file("Thyroid", path),4).iloc[0,4], 1.2626262626262625)       

    def get_subpanel_coverage(self):
        self.assertEqual(len(get_subpanel_coverage("Breast", path)),1)
        self.assertEqual(len(get_subpanel_coverage("Colorectal", path)),1)
        self.assertEqual(len(get_subpanel_coverage("GIST", path)),1)
        self.assertEqual(len(get_subpanel_coverage("Glioma", path)),1)
        self.assertEqual(len(get_subpanel_coverage("HN", path)),1)
        self.assertEqual(len(get_subpanel_coverage("Lung", path)),1)
        self.assertEqual(len(get_subpanel_coverage("Melanoma", path)),1)
        self.assertEqual(len(get_subpanel_coverage("Ovarian", path)),1)
        self.assertEqual(len(get_subpanel_coverage("Prostate", path)),1)
        self.assertEqual(len(get_subpanel_coverage("Thyroid", path)),1)


    def test_match_polys_and_artefacts(self):
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("Breast", path, "tester"), get_variantReport_NTC("Breast", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("Breast", path) ,get_variant_report("Breast", path, "tester")))).iloc[0,18],"")
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("Glioma", path, "tester"), get_variantReport_NTC("Glioma", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("Glioma", path) ,get_variant_report("Glioma", path, "tester")))).iloc[1,16],"NO")
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("Glioma", path, "tester"), get_variantReport_NTC("Glioma", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("Glioma", path) ,get_variant_report("Glioma", path, "tester")))).iloc[1,18],"")
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("HN", path, "tester"), get_variantReport_NTC("HN", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("HN", path) ,get_variant_report("HN", path, "tester")))).iloc[1,16],"NO")
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("HN", path, "tester"), get_variantReport_NTC("HN", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("HN", path) ,get_variant_report("HN", path, "tester")))).iloc[1,18],"")
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("Lung", path, "tester"), get_variantReport_NTC("Lung", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("Lung", path) ,get_variant_report("Lung", path, "tester")))).iloc[1,16],"YES")
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("Lung", path, "tester"), get_variantReport_NTC("Lung", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("Lung", path) ,get_variant_report("Lung", path, "tester")))).iloc[1,18],325.728)
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("Lung", path, "tester"), get_variantReport_NTC("Lung", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("Lung", path) ,get_variant_report("Lung", path, "tester")))).iloc[0,19],0.42280285035629456)
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("Melanoma", path, "tester"), get_variantReport_NTC("Melanoma", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("Melanoma", path) ,get_variant_report("Melanoma", path, "tester")))).iloc[0,16],"NO")
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("Melanoma", path, "tester"), get_variantReport_NTC("Melanoma", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("Melanoma", path) ,get_variant_report("Melanoma", path, "tester")))).iloc[0,18],"")
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("Ovarian", path, "tester"), get_variantReport_NTC("Ovarian", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("Ovarian", path) ,get_variant_report("Ovarian", path, "tester")))).iloc[1,18],"")
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("Prostate", path, "tester"), get_variantReport_NTC("Prostate", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("Prostate", path) ,get_variant_report("Prostate", path, "tester")))).iloc[0,16],"NO")
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("Prostate", path, "tester"), get_variantReport_NTC("Prostate", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("Prostate", path) ,get_variant_report("Prostate", path, "tester")))).iloc[0,18],"")
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("Thyroid", path, "tester"), get_variantReport_NTC("Thyroid", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("Thyroid", path) ,get_variant_report("Thyroid", path, "tester")))).iloc[6,16], "YES" )
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("Thyroid", path, "tester"), get_variantReport_NTC("Thyroid", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("Thyroid", path) ,get_variant_report("Thyroid", path, "tester")))).iloc[6,18], 21.675 )
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("Thyroid", path, "tester"), get_variantReport_NTC("Thyroid", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("Thyroid", path) ,get_variant_report("Thyroid", path, "tester")))).iloc[6,19],0.09980430528375735)

if __name__ == '__main__':
    unittest.main()
