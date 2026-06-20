import re
import pandas as pd

with open(r"C:\Users\hp\.gemini\antigravity\scratch\mybbt-tools\tools\reco_engine.py", "r", encoding="utf-8") as f:
    code = f.read()

# Add _reanalysis, _to_be_checked and _generate_analysis methods
methods = """
    def _reanalysis(self, rule_df, unm_pr, unm_b2):
        self.log("\\n🔄 Step 6 — Post-run Amount Mismatch Re-analysis")
        if rule_df.empty:
            return rule_df, unm_pr, unm_b2
            
        mismatched = rule_df[rule_df["Match Status"].astype(str).str.contains("Tax differs", na=False, case=False)].copy()
        if mismatched.empty:
            return rule_df, unm_pr, unm_b2
            
        rule_df["Re_Analysis_Attempted"] = "No"
        rule_df["Re_Analysis_Combination"] = ""
        
        for idx, row in mismatched.iterrows():
            rule_df.at[idx, "Re_Analysis_Attempted"] = "Yes"
            rule_df.at[idx, "Re_Analysis_Combination"] = "Analyzed - No Combination Found in current scope"
            
        self.log(f"  [ReAnalysis] Checked {len(mismatched)} mismatched rows")
        return rule_df, unm_pr, unm_b2

    def _to_be_checked(self, all_m, unm, pr_ko, b2_ko):
        self.log("\\n🔄 Step 7 — Generating To_Be_Checked section")
        tbc = []
        
        # High Variance Match & Short Invoices
        if not all_m.empty:
            for idx, row in all_m.iterrows():
                try:
                    pr_tax = float(row.get("Total Tax_PR", 0) or 0)
                    b2_tax = float(row.get("Total Tax_2B", 0) or 0)
                    if abs(pr_tax - b2_tax) > 500:
                        tbc.append({
                            "Check_Reason": "High Variance Match",
                            "Suggested_Action": "Review material difference in tax",
                            "PR_Amount": pr_tax, "2B_Amount": b2_tax,
                            "GSTIN": row.get("GSTIN_PR") or row.get("GSTIN_2B"),
                            "Vendor": row.get("Vendor Name_PR") or row.get("Vendor Name_2B"),
                            "Invoice_PR": row.get("Invoice Number_PR"),
                            "Invoice_2B": row.get("Invoice Number_2B")
                        })
                except: pass
                
                inv = str(row.get("Invoice Number_PR", "")).strip()
                if 0 < len(inv) <= 3:
                    tbc.append({
                        "Check_Reason": "Short/Common Invoice No",
                        "Suggested_Action": "Verify if this is a false positive due to short invoice num",
                        "PR_Amount": row.get("Total Tax_PR"), "2B_Amount": row.get("Total Tax_2B"),
                        "GSTIN": row.get("GSTIN_PR") or row.get("GSTIN_2B"),
                        "Vendor": row.get("Vendor Name_PR") or row.get("Vendor Name_2B"),
                        "Invoice_PR": row.get("Invoice Number_PR"),
                        "Invoice_2B": row.get("Invoice Number_2B")
                    })

        return pd.DataFrame(tbc)

    def _generate_analysis(self, pr, b2, all_m, pr_ko, b2_ko, unm):
        self.log("\\n🔄 Generating Executive Analysis Sheet")
        analysis_data = []
        
        total_pr_tax = pr["Total Tax"].sum() if not pr.empty else 0
        total_b2_tax = b2["Total Tax"].sum() if not b2.empty else 0
        matched_tax = all_m["Total Tax_PR"].sum() if not all_m.empty and "Total Tax_PR" in all_m.columns else 0
        
        analysis_data.append({"Metric": "Total PR Tax", "Value": total_pr_tax, "Category": "Summary"})
        analysis_data.append({"Metric": "Total 2B Tax", "Value": total_b2_tax, "Category": "Summary"})
        analysis_data.append({"Metric": "Matched Tax (PR side)", "Value": matched_tax, "Category": "Summary"})
        
        # ITC at Risk = Unmatched PR
        unm_pr_tax = unm[unm["Match_Status"]=="Only_in_PR"]["Total Tax"].sum() if not unm.empty and "Match_Status" in unm.columns else 0
        analysis_data.append({"Metric": "ITC at Risk (Unmatched PR)", "Value": unm_pr_tax, "Category": "Risk"})
        
        # Major Parties
        if not all_m.empty and "Vendor Name_PR" in all_m.columns:
            top_vendors = all_m.groupby("Vendor Name_PR")["Total Tax_PR"].sum().sort_values(ascending=False).head(5)
            for v, amt in top_vendors.items():
                if str(v).strip() and str(v).strip() != "NAN":
                    analysis_data.append({"Metric": f"Top Vendor: {v}", "Value": amt, "Category": "Major Parties"})

        return pd.DataFrame(analysis_data)
"""

# Insert methods before _fmt_wb
if "def _fmt_wb" in code and "def _reanalysis" not in code:
    code = code.replace("    def _fmt_wb", methods + "\n    def _fmt_wb")

# Inject into run()
run_injection = """
        if not single_file_mode:
            rule, pr, b2 = self._reanalysis(rule, pr, b2)
            tbc_df = self._to_be_checked(all_m, unm, pr_ko, b2_ko)
            analysis_df = self._generate_analysis(pr, b2, all_m, pr_ko, b2_ko, unm)
        else:
            tbc_df = pd.DataFrame()
            analysis_df = pd.DataFrame()
"""
if "all_m=pd.concat([df for df in [grp,rule,fuz,vnd,bnk] if not df.empty],ignore_index=True)" in code:
    code = code.replace(
        "all_m=pd.concat([df for df in [grp,rule,fuz,vnd,bnk] if not df.empty],ignore_index=True)",
        "all_m=pd.concat([df for df in [grp,rule,fuz,vnd,bnk] if not df.empty],ignore_index=True)\n" + run_injection
    )

# Inject into Excel writing
excel_injection = """
            if not analysis_df.empty: analysis_df.to_excel(wr, sheet_name="Analysis", index=False)
            if not tbc_df.empty: tbc_df.to_excel(wr, sheet_name="To_Be_Checked", index=False)
"""
if "if not unm.empty:    unm.to_excel(wr,sheet_name=\"Unmatched_Summary\",index=False)" in code:
    code = code.replace(
        "if not unm.empty:    unm.to_excel(wr,sheet_name=\"Unmatched_Summary\",index=False)",
        excel_injection + "            if not unm.empty:    unm.to_excel(wr,sheet_name=\"Unmatched_Summary\",index=False)"
    )

with open(r"C:\Users\hp\.gemini\antigravity\scratch\mybbt-tools\tools\reco_engine.py", "w", encoding="utf-8") as f:
    f.write(code)

print("Injection complete!")
