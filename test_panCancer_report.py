import unittest

from panCancer_report import *

path="./tests/"

coverage_value="250x"

class test_virtualhood(unittest.TestCase):

    def test_get_variantReport_NTC(self):
        self.assertEqual(len(get_variantReport_NTC("BRAF", path)),1)
        self.assertEqual(len(get_variantReport_NTC("Breast", path)),2)
        self.assertEqual(len(get_variantReport_NTC("Colorectal", path)),1)
        self.assertEqual(len(get_variantReport_NTC("GIST", path)),8)
        self.assertEqual(len(get_variantReport_NTC("Glioma", path)),7)
        self.assertEqual(len(get_variantReport_NTC("HeadAndNeck", path)),6)
        self.assertEqual(len(get_variantReport_NTC("Lung", path)),5)
        self.assertEqual(len(get_variantReport_NTC("Melanoma", path)),4)
        self.assertEqual(len(get_variantReport_NTC("Ovarian", path)),3)
        self.assertEqual(len(get_variantReport_NTC("Prostate", path)),2)
        self.assertEqual(len(get_variantReport_NTC("Thyroid", path)),2)

    def test_get_variant_report(self):
        self.assertEqual(len(get_variant_report("BRAF", path, "tester")),2)
        self.assertEqual(len(get_variant_report("Breast", path, "tester")),2)
        self.assertEqual(len(get_variant_report("Colorectal", path, "tester")),0)
        self.assertEqual(len(get_variant_report("GIST", path, "tester")),0)
        self.assertEqual(len(get_variant_report("Glioma", path, "tester")),3)
        self.assertEqual(len(get_variant_report("HeadAndNeck", path, "tester")),4)
        self.assertEqual(len(get_variant_report("Lung", path, "tester")),6)
        self.assertEqual(len(get_variant_report("Melanoma", path, "tester")),3)
        self.assertEqual(len(get_variant_report("Ovarian", path, "tester")),5)
        self.assertEqual(len(get_variant_report("Prostate", path, "tester")),1)
        self.assertEqual(len(get_variant_report("Thyroid", path, "tester")),7)

    def test_add_extra_columns_NTC_report(self):
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("BRAF", path) ,get_variant_report("BRAF", path, "tester")).iloc[0,11],9.00)
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Breast", path) ,get_variant_report("Breast", path, "tester")).iloc[1,10],"NO")
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Breast", path) ,get_variant_report("Breast", path, "tester")).iloc[0,11],9.00)
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Colorectal", path) ,get_variant_report("Colorectal", path, "tester")).iloc[0,10],"NO")
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Colorectal", path), get_variant_report("Colorectal", path, "tester")).iloc[0,11],5.00)
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("GIST", path) ,get_variant_report("GIST", path, "tester")).iloc[3,10],"NO")
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("GIST", path), get_variant_report("GIST", path, "tester")).iloc[4,11],325.7280)
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Glioma", path) ,get_variant_report("Glioma", path, "tester")).iloc[3,10],"NO")
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Glioma", path), get_variant_report("Glioma", path, "tester")).iloc[4,11],325.7280)        
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("HeadAndNeck", path) ,get_variant_report("HeadAndNeck", path, "tester")).iloc[5,10],"NO")
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("HeadAndNeck", path), get_variant_report("HeadAndNeck", path, "tester")).iloc[1,11],331.1184)
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Lung", path) ,get_variant_report("Lung", path, "tester")).iloc[3,10],"YES")
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Lung", path), get_variant_report("Lung", path, "tester")).iloc[4,11],325.7280)  
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Melanoma", path) ,get_variant_report("Melanoma", path, "tester")).iloc[3,10],"NO")
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Melanoma", path), get_variant_report("Melanoma", path, "tester")).iloc[3,11],239.3388)
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Prostate", path) ,get_variant_report("Prostate", path, "tester")).iloc[0,10],"NO")
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Prostate", path), get_variant_report("Prostate", path, "tester")).iloc[1,11],331.1184)
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Thyroid", path) ,get_variant_report("Thyroid", path, "tester")).iloc[0,10],"NO")
        self.assertEqual(add_extra_columns_NTC_report(get_variantReport_NTC("Thyroid", path), get_variant_report("Thyroid", path, "tester")).iloc[0,11],14.0)
   
    def test_expand_variant_report(self):
        self.assertEqual(expand_variant_report(get_variant_report("Breast", path, "tester"), get_variantReport_NTC("Breast", path)).iloc[0,15],0)
        self.assertEqual(expand_variant_report(get_variant_report("Glioma", path, "tester"), get_variantReport_NTC("Glioma", path)).iloc[0,15],0)
        self.assertEqual(expand_variant_report(get_variant_report("Lung", path, "tester"), get_variantReport_NTC("Lung", path)).iloc[2,15],'195.8125')    
        self.assertEqual(expand_variant_report(get_variant_report("Melanoma", path, "tester"), get_variantReport_NTC("Melanoma", path)).iloc[0,15],0)
        self.assertEqual(expand_variant_report(get_variant_report("Ovarian", path, "tester"), get_variantReport_NTC("Ovarian", path)).iloc[0,15],0)
        self.assertEqual(expand_variant_report(get_variant_report("Prostate", path, "tester"), get_variantReport_NTC("Prostate", path)).iloc[0,15],0) 

    def test_get_gaps_file(self):
        self.assertEqual(len(get_gaps_file("BRAF", path, "tester", coverage_value)),8)
        self.assertEqual(len(get_gaps_file("Breast", path, "tester", coverage_value)),14)
        self.assertEqual(len(get_gaps_file("Colorectal", path, "tester", coverage_value)),13)
        self.assertEqual(len(get_gaps_file("GIST", path, "tester", coverage_value)),12)
        self.assertEqual(len(get_gaps_file("Glioma", path, "tester", coverage_value)),11)
        self.assertEqual(len(get_gaps_file("HeadAndNeck", path, "tester", coverage_value)),10)
        self.assertEqual(len(get_gaps_file("Lung", path, "tester", coverage_value)),9)
        self.assertEqual(len(get_gaps_file("Melanoma", path, "tester", coverage_value)),8)
        self.assertEqual(len(get_gaps_file("Ovarian", path, "tester", coverage_value)),7)
        self.assertEqual(len(get_gaps_file("Prostate", path, "tester", coverage_value)),6)
        self.assertEqual(len(get_gaps_file("Thyroid", path, "tester", coverage_value)),5)        

    def test_get_CNV_file(self):
        self.assertEqual(len(get_CNV_file("Breast", path, "tester")),5)
        self.assertEqual(len(get_CNV_file("Colorectal", path, "tester")),3)
        self.assertEqual(len(get_CNV_file("GIST", path, "tester")),1)
        self.assertEqual(len(get_CNV_file("Glioma", path, "tester")),3)
        self.assertEqual(len(get_CNV_file("HeadAndNeck", path, "tester")),4)
        self.assertEqual(len(get_CNV_file("Lung", path, "tester")),6)
        self.assertEqual(len(get_CNV_file("Melanoma", path, "tester")),0)
        self.assertEqual(len(get_CNV_file("Ovarian", path, "tester")),7)
        self.assertEqual(len(get_CNV_file("Prostate", path, "tester")),3)
        self.assertEqual(len(get_CNV_file("Thyroid", path, "tester")),1)


    def test_get_hotspots_coverage_file(self):
        self.assertEqual(len(get_hotspots_coverage_file("BRAF", path, "tester", coverage_value)),7)
        self.assertEqual(len(get_hotspots_coverage_file("Breast", path, "tester", coverage_value)),8)
        self.assertEqual(len(get_hotspots_coverage_file("Colorectal", path, "tester", coverage_value)),7)
        self.assertEqual(len(get_hotspots_coverage_file("GIST", path, "tester", coverage_value)),6)
        self.assertEqual(len(get_hotspots_coverage_file("Glioma", path, "tester", coverage_value)),5)
        self.assertEqual(len(get_hotspots_coverage_file("HeadAndNeck", path, "tester", coverage_value)),4)
        self.assertEqual(len(get_hotspots_coverage_file("Lung", path, "tester", coverage_value)),3)
        self.assertEqual(len(get_hotspots_coverage_file("Melanoma", path, "tester", coverage_value)),2)
        self.assertEqual(len(get_hotspots_coverage_file("Ovarian", path, "tester", coverage_value)),1)
        self.assertEqual(len(get_hotspots_coverage_file("Prostate", path, "tester", coverage_value)),0)
        self.assertEqual(len(get_hotspots_coverage_file("Thyroid", path, "tester", coverage_value)),4)


    def test_get_NTC_hotspots_coverage_file(self):
        self.assertEqual(len(get_NTC_hotspots_coverage_file("BRAF", path, coverage_value)),8)
        self.assertEqual(len(get_NTC_hotspots_coverage_file("Breast", path, coverage_value)),8)
        self.assertEqual(len(get_NTC_hotspots_coverage_file("Colorectal", path, coverage_value)),7)
        self.assertEqual(len(get_NTC_hotspots_coverage_file("GIST", path, coverage_value)),6)
        self.assertEqual(len(get_NTC_hotspots_coverage_file("Glioma", path, coverage_value)),5)
        self.assertEqual(len(get_NTC_hotspots_coverage_file("HeadAndNeck", path, coverage_value)),4)
        self.assertEqual(len(get_NTC_hotspots_coverage_file("Lung", path, coverage_value)),3)
        self.assertEqual(len(get_NTC_hotspots_coverage_file("Melanoma", path, coverage_value)),2)
        self.assertEqual(len(get_NTC_hotspots_coverage_file("Ovarian", path, coverage_value)),1)
        self.assertEqual(len(get_NTC_hotspots_coverage_file("Prostate", path, coverage_value)),0)
        self.assertEqual(len(get_NTC_hotspots_coverage_file("Thyroid", path, coverage_value)),4)

   
    def test_add_columns_hotspots_coverage(self):
        self.assertEqual(add_columns_hotspots_coverage(get_hotspots_coverage_file("Breast", path, "tester", coverage_value), get_NTC_hotspots_coverage_file("Breast", path, coverage_value),path, "tester", "Breast")[0].iloc[2,5], 5.818181818181818)
        self.assertEqual(add_columns_hotspots_coverage(get_hotspots_coverage_file("Colorectal", path, "tester", coverage_value), get_NTC_hotspots_coverage_file("Colorectal", path,coverage_value), path,"tester", "Colorectal")[0].iloc[2,5], 34.18181818181818)
        self.assertEqual(add_columns_hotspots_coverage(get_hotspots_coverage_file("GIST", path, "tester", coverage_value), get_NTC_hotspots_coverage_file("GIST", path, coverage_value), path, "tester", "GIST")[0].iloc[2,5], 25.818181818181817)
        self.assertEqual(add_columns_hotspots_coverage(get_hotspots_coverage_file("Glioma", path, "tester", coverage_value), get_NTC_hotspots_coverage_file("Glioma", path, coverage_value), path, "tester", "Glioma")[0].iloc[1,5], 31.026993484331367)
        self.assertEqual(add_columns_hotspots_coverage(get_hotspots_coverage_file("HeadAndNeck", path, "tester", coverage_value), get_NTC_hotspots_coverage_file("HeadAndNeck", path, coverage_value), path, "tester", "HeadAndNeck")[0].iloc[1,5], 6.211180124223603)
        self.assertEqual(add_columns_hotspots_coverage(get_hotspots_coverage_file("Lung", path, "tester", coverage_value), get_NTC_hotspots_coverage_file("Lung", path, coverage_value), path, "tester", "Lung")[0].iloc[1,5], 2.1739130434782608)
        self.assertEqual(add_columns_hotspots_coverage(get_hotspots_coverage_file("Melanoma", path, "tester", coverage_value), get_NTC_hotspots_coverage_file("Melanoma", path, coverage_value), path, "tester", "Melanoma")[0].iloc[1,5],2.1739130434782608)
        self.assertEqual(add_columns_hotspots_coverage(get_hotspots_coverage_file("Ovarian", path, "tester", coverage_value), get_NTC_hotspots_coverage_file("Ovarian", path, coverage_value), path, "tester", "Ovarian")[0].iloc[0,5],2.0202020202020203)
        self.assertEqual(add_columns_hotspots_coverage(get_hotspots_coverage_file("Thyroid", path, "tester", coverage_value), get_NTC_hotspots_coverage_file("Thyroid", path, coverage_value), path, "tester", "Thyroid")[0].iloc[3,5],5.105105105105105)

    def test_get_genescreen_coverage_file(self):
        self.assertEqual(len(get_genescreen_coverage_file("Breast", path, "tester", coverage_value)),10)
        self.assertEqual(len(get_genescreen_coverage_file("Colorectal", path, "tester", coverage_value)),8)
        self.assertEqual(len(get_genescreen_coverage_file("GIST", path, "tester", coverage_value)),8)
        self.assertEqual(len(get_genescreen_coverage_file("Glioma", path, "tester", coverage_value)),7)
        self.assertEqual(len(get_genescreen_coverage_file("HeadAndNeck", path, "tester", coverage_value)),6)
        self.assertEqual(len(get_genescreen_coverage_file("Lung", path, "tester", coverage_value)),5)
        self.assertEqual(len(get_genescreen_coverage_file("Melanoma", path, "tester", coverage_value)),4)
        self.assertEqual(len(get_genescreen_coverage_file("Ovarian", path, "tester", coverage_value)),3)
        self.assertEqual(len(get_genescreen_coverage_file("Prostate", path, "tester", coverage_value)),2)
        self.assertEqual(len(get_genescreen_coverage_file("Thyroid", path, "tester", coverage_value)),1)

    def test_get_NTC_genescreen_coverage_file(self):
        self.assertEqual(len(get_NTC_genescreen_coverage_file("Breast", path, coverage_value)),10)
        self.assertEqual(len(get_NTC_genescreen_coverage_file("Colorectal", path, coverage_value)),8)
        self.assertEqual(len(get_NTC_genescreen_coverage_file("GIST", path, coverage_value)),8)
        self.assertEqual(len(get_NTC_genescreen_coverage_file("Glioma", path, coverage_value)),7)
        self.assertEqual(len(get_NTC_genescreen_coverage_file("HeadAndNeck", path, coverage_value)),6)
        self.assertEqual(len(get_NTC_genescreen_coverage_file("Lung", path, coverage_value)),6)
        self.assertEqual(len(get_NTC_genescreen_coverage_file("Melanoma", path, coverage_value)),6)
        self.assertEqual(len(get_NTC_genescreen_coverage_file("Ovarian", path, coverage_value)),5)
        self.assertEqual(len(get_NTC_genescreen_coverage_file("Prostate", path, coverage_value)),4)
        self.assertEqual(len(get_NTC_genescreen_coverage_file("Thyroid", path, coverage_value)),3)

    def test_add_columns_genescreen_coverage_file(self):
        self.assertEqual(add_columns_genescreen_coverage(get_genescreen_coverage_file("Breast", path, "tester", coverage_value), get_NTC_genescreen_coverage_file("Breast", path, coverage_value),8, path, "tester", "Breast").iloc[2,4], 275)
        self.assertEqual(add_columns_genescreen_coverage(get_genescreen_coverage_file("Colorectal", path, "tester", coverage_value), get_NTC_genescreen_coverage_file("Colorectal", path, coverage_value),7, path, "tester", "Colorectal").iloc[7,4], 135)
        self.assertEqual(add_columns_genescreen_coverage(get_genescreen_coverage_file("GIST", path, "tester", coverage_value), get_NTC_genescreen_coverage_file("GIST", path, coverage_value),6, path, "tester", "GIST").iloc[4,4], 15)        
        self.assertEqual(add_columns_genescreen_coverage(get_genescreen_coverage_file("Glioma", path, "tester", coverage_value), get_NTC_genescreen_coverage_file("Glioma", path, coverage_value),5, path, "tester", "Glioma").iloc[4,4], 1)
        self.assertEqual(add_columns_genescreen_coverage(get_genescreen_coverage_file("HeadAndNeck", path, "tester", coverage_value), get_NTC_genescreen_coverage_file("HeadAndNeck", path, coverage_value),4, path, "tester", "HeadAndNeck").iloc[4,4], 3) 
        self.assertEqual(add_columns_genescreen_coverage(get_genescreen_coverage_file("Lung", path, "tester", coverage_value), get_NTC_genescreen_coverage_file("Lung", path, coverage_value),3, path, "tester", "Lung").iloc[4,4], 8)        
        self.assertEqual(add_columns_genescreen_coverage(get_genescreen_coverage_file("Melanoma", path, "tester",coverage_value), get_NTC_genescreen_coverage_file("Melanoma", path, coverage_value),2, path, "tester", "Melanoma").iloc[3,4], 18)
        self.assertEqual(add_columns_genescreen_coverage(get_genescreen_coverage_file("Ovarian", path, "tester", coverage_value), get_NTC_genescreen_coverage_file("Ovarian", path, coverage_value),1, path, "tester", "Ovarian").iloc[2,4], 1)        
        self.assertEqual(add_columns_genescreen_coverage(get_genescreen_coverage_file("Prostate", path, "tester", coverage_value), get_NTC_genescreen_coverage_file("Prostate", path, coverage_value),0, path, "tester", "Prostate").iloc[1,4], 52)        
        self.assertEqual(add_columns_genescreen_coverage(get_genescreen_coverage_file("Thyroid", path, "tester", coverage_value), get_NTC_genescreen_coverage_file("Thyroid", path, coverage_value),4, path, "tester", "Thyroid").iloc[0,4], 5)       

    def test_get_subpanel_coverage(self):
        self.assertEqual(len(get_subpanel_coverage("BRAF", path,"tester", coverage_value)),1)
        self.assertEqual(len(get_subpanel_coverage("Breast", path,"tester", coverage_value)),1)
        self.assertEqual(len(get_subpanel_coverage("Colorectal", path,"tester", coverage_value)),1)
        self.assertEqual(len(get_subpanel_coverage("GIST", path, "tester", coverage_value)),1)
        self.assertEqual(len(get_subpanel_coverage("Glioma", path, "tester", coverage_value)),1)
        self.assertEqual(len(get_subpanel_coverage("HeadAndNeck", path,"tester", coverage_value)),1)
        self.assertEqual(len(get_subpanel_coverage("Lung", path,"tester", coverage_value)),1)
        self.assertEqual(len(get_subpanel_coverage("Melanoma", path, "tester",coverage_value)),1)
        self.assertEqual(len(get_subpanel_coverage("Ovarian", path, "tester",coverage_value)),1)
        self.assertEqual(len(get_subpanel_coverage("Prostate", path,"tester", coverage_value)),1)
        self.assertEqual(len(get_subpanel_coverage("Thyroid", path, "tester", coverage_value)),1)


    def test_match_polys_and_artefacts(self):
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("Breast", path, "tester"), get_variantReport_NTC("Breast", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("Breast", path) ,get_variant_report("Breast", path, "tester")))).iloc[0,18],"")
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("Glioma", path, "tester"), get_variantReport_NTC("Glioma", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("Glioma", path) ,get_variant_report("Glioma", path, "tester")))).iloc[1,16],"NO")
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("Glioma", path, "tester"), get_variantReport_NTC("Glioma", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("Glioma", path) ,get_variant_report("Glioma", path, "tester")))).iloc[1,18],"")
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("HeadAndNeck", path, "tester"), get_variantReport_NTC("HeadAndNeck", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("HeadAndNeck", path) ,get_variant_report("HeadAndNeck", path, "tester")))).iloc[1,16],"NO")
        self.assertEqual(match_polys_and_artefacts((expand_variant_report(get_variant_report("HeadAndNeck", path, "tester"), get_variantReport_NTC("HeadAndNeck", path))), (add_extra_columns_NTC_report(get_variantReport_NTC("HeadAndNeck", path) ,get_variant_report("HeadAndNeck", path, "tester")))).iloc[1,18],"")
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
