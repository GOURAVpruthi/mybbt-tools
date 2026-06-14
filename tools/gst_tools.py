"""
GST Tools — Web wrapper around gst_engine.py (user's original code).

Exposes all processing functions through a clean class API for Flask.
The original engine (gst_engine.py) is NOT modified — all logic is preserved exactly.
"""

import os
import sys
import tempfile
import shutil
from datetime import datetime
from pathlib import Path

# Ensure tools/ dir is on sys.path so gst_engine can be found
_TOOLS_DIR = os.path.dirname(os.path.abspath(__file__))
if _TOOLS_DIR not in sys.path:
    sys.path.insert(0, _TOOLS_DIR)

# ─── Lazy engine loader ───────────────────────────────────────────────────────
_engine = None

def _get_engine():
    """Load gst_engine once, lazily (avoids slow startup on first import)."""
    global _engine
    if _engine is None:
        import gst_engine as _e
        # Ensure base64-embedded GSTR-1 & GSTR-3B engines are decoded
        _e._load_extractor_engines()
        _engine = _e
    return _engine


# ─── Timestamp helper ─────────────────────────────────────────────────────────
def _ts():
    return datetime.now().strftime('%Y%m%d_%H%M%S')


# ═════════════════════════════════════════════════════════════════════════════
#  GSTTools class — called by Flask routes in app.py
# ═════════════════════════════════════════════════════════════════════════════
class GSTTools:
    def __init__(self, upload_folder, output_folder):
        self.upload_folder = upload_folder
        self.output_folder = output_folder

    # ── GSTR-1: Multiple PDFs → Consolidated Excel ───────────────────────────
    def process_gstr1(self, input_paths):
        """
        Process GSTR-1 PDFs downloaded from GST portal.
        input_paths: list of uploaded PDF file paths
        Returns: dict with success, filename, stats
        """
        try:
            e = _get_engine()
            # Create a temp folder with the uploaded PDFs (engine needs a folder)
            tmp_dir = Path(self.upload_folder) / f"gstr1_tmp_{_ts()}"
            tmp_dir.mkdir(exist_ok=True)

            for p in input_paths:
                shutil.copy2(p, tmp_dir / Path(p).name)

            out_name = f"GSTR1_Consolidated_{_ts()}.xlsx"
            out_path = os.path.join(self.output_folder, out_name)

            logs = []
            def _log(msg, level='info'):
                logs.append(msg)

            ok, fail, total, _ = e.process_pdfs(
                tmp_dir, Path(out_path),
                on_progress=None,
                on_log=_log
            )

            # Cleanup temp dir
            shutil.rmtree(tmp_dir, ignore_errors=True)

            if ok == 0 and fail > 0:
                return {
                    'success': False,
                    'error': f'All {fail} PDF(s) failed to process. Check that files are GSTR-1 PDFs from GST portal.',
                    'logs': logs
                }

            return {
                'success': True,
                'filename': out_name,
                'ok': ok,
                'fail': fail,
                'total': total,
                'logs': logs,
                'message': f'✅ Processed {ok}/{total} PDFs successfully. Excel report ready.'
            }

        except Exception as ex:
            return {'success': False, 'error': str(ex)}

    # ── GSTR-2B: Multiple Excel files → Consolidated ─────────────────────────
    def process_gstr2b(self, input_paths):
        """
        Process GSTR-2B Excel files downloaded from GST portal.
        input_paths: list of uploaded .xlsx file paths
        """
        try:
            e = _get_engine()
            all_data = []
            errors = []

            for p in input_paths:
                try:
                    data = e.parse_gstr2b_file(p)
                    all_data.append(data)
                except Exception as ex:
                    errors.append(f"{Path(p).name}: {ex}")

            if not all_data:
                return {'success': False, 'error': 'No valid GSTR-2B files found. ' + '; '.join(errors)}

            out_name = f"GSTR2B_Consolidated_{_ts()}.xlsx"
            out_path = os.path.join(self.output_folder, out_name)
            e.write_consolidated_gstr2b(all_data, out_path)

            return {
                'success': True,
                'filename': out_name,
                'files_processed': len(all_data),
                'errors': errors,
                'message': f'✅ {len(all_data)} GSTR-2B file(s) consolidated successfully.'
            }

        except Exception as ex:
            return {'success': False, 'error': str(ex)}

    # ── GSTR-3B: Multiple PDFs → Consolidated Excel ──────────────────────────
    def process_gstr3b(self, input_paths):
        """
        Process GSTR-3B PDFs from GST portal.
        """
        try:
            e = _get_engine()
            # GSTR-3B engine is in _GSTR3B_NS namespace
            ns = e._GSTR3B_NS
            all_data = []
            errors = []

            for p in input_paths:
                try:
                    meta, rows = ns['parse_pdf'](p)
                    all_data.append({
                        'meta': meta, 'rows': rows,
                        'source_file': Path(p).name,
                        'status': 'OK', 'notes': ''
                    })
                except Exception as ex:
                    errors.append(f"{Path(p).name}: {ex}")

            if not all_data:
                return {'success': False, 'error': 'No valid GSTR-3B PDFs found. ' + '; '.join(errors)}

            out_name = f"GSTR3B_Consolidated_{_ts()}.xlsx"
            out_path = os.path.join(self.output_folder, out_name)
            ns['write_excel'](all_data, out_path)

            return {
                'success': True,
                'filename': out_name,
                'files_processed': len(all_data),
                'errors': errors,
                'message': f'✅ {len(all_data)} GSTR-3B PDF(s) consolidated successfully.'
            }

        except Exception as ex:
            return {'success': False, 'error': str(ex)}

    # ── GSTR-9/9C: Multiple PDFs → Consolidated Excel ────────────────────────
    def process_gstr9(self, input_paths):
        """
        Process GSTR-9 / 9C Annual Return PDFs from GST portal.
        """
        try:
            e = _get_engine()
            all_data = []
            errors = []

            for p in input_paths:
                try:
                    data = e.parse_gstr9_or_9c_pdf(p)
                    all_data.append(data)
                except Exception as ex:
                    errors.append(f"{Path(p).name}: {ex}")

            if not all_data:
                return {'success': False, 'error': 'No valid GSTR-9/9C PDFs found. ' + '; '.join(errors)}

            out_name = f"GSTR9_Consolidated_{_ts()}.xlsx"
            out_path = os.path.join(self.output_folder, out_name)
            e.write_consolidated_gstr9_9c(all_data, out_path)

            return {
                'success': True,
                'filename': out_name,
                'files_processed': len(all_data),
                'errors': errors,
                'message': f'✅ {len(all_data)} GSTR-9/9C PDF(s) consolidated successfully.'
            }

        except Exception as ex:
            return {'success': False, 'error': str(ex)}

    # ── Tax Comparison (GSTR-1 vs 3B): Excel → Consolidated ─────────────────
    def process_tax_comparison(self, input_paths, mode='all'):
        """
        Process GSTN Tax Liability & ITC Comparison Excel files.
        mode: 'all' / 'liability_only' / 'itc_only'
        """
        try:
            e = _get_engine()
            all_data = []
            errors = []

            for p in input_paths:
                try:
                    data = e.parse_tax_comparison_file(p)
                    all_data.append(data)
                except Exception as ex:
                    errors.append(f"{Path(p).name}: {ex}")

            if not all_data:
                return {'success': False, 'error': 'No valid comparison files found. ' + '; '.join(errors)}

            out_name = f"TaxComparison_Consolidated_{_ts()}.xlsx"
            out_path = os.path.join(self.output_folder, out_name)
            e.write_consolidated_comparison(all_data, out_path, mode)

            return {
                'success': True,
                'filename': out_name,
                'files_processed': len(all_data),
                'errors': errors,
                'message': f'✅ {len(all_data)} comparison file(s) consolidated.'
            }

        except Exception as ex:
            return {'success': False, 'error': str(ex)}

    # ── ECRRS (Electronic Credit Reversal): CSV → Consolidated Excel ──────────
    def process_ecrrs(self, input_paths):
        """Process ECRRS CSV files from GST portal."""
        try:
            e = _get_engine()
            all_data = []
            errors = []

            for p in input_paths:
                try:
                    data = e.parse_ecrrs_csv(p)
                    all_data.append(data)
                except Exception as ex:
                    errors.append(f"{Path(p).name}: {ex}")

            if not all_data:
                return {'success': False, 'error': 'No valid ECRRS files. ' + '; '.join(errors)}

            out_name = f"ECRRS_Consolidated_{_ts()}.xlsx"
            out_path = os.path.join(self.output_folder, out_name)
            e.write_consolidated_ecrrs(all_data, out_path)

            return {
                'success': True,
                'filename': out_name,
                'files_processed': len(all_data),
                'message': f'✅ {len(all_data)} ECRRS file(s) processed.'
            }

        except Exception as ex:
            return {'success': False, 'error': str(ex)}

    # ── Electronic Credit Ledger: CSV → Consolidated Excel ───────────────────
    def process_credit_ledger(self, input_paths):
        """Process Electronic Credit Ledger CSVs from GST portal."""
        try:
            e = _get_engine()
            all_data = []
            errors = []

            for p in input_paths:
                try:
                    data = e.parse_ecl_csv(p)
                    all_data.append(data)
                except Exception as ex:
                    errors.append(f"{Path(p).name}: {ex}")

            if not all_data:
                return {'success': False, 'error': 'No valid Credit Ledger files. ' + '; '.join(errors)}

            out_name = f"CreditLedger_Consolidated_{_ts()}.xlsx"
            out_path = os.path.join(self.output_folder, out_name)
            e.write_consolidated_ecl(all_data, out_path)

            return {
                'success': True,
                'filename': out_name,
                'files_processed': len(all_data),
                'message': f'✅ {len(all_data)} Credit Ledger file(s) processed.'
            }

        except Exception as ex:
            return {'success': False, 'error': str(ex)}

    # ── Electronic Cash Ledger: CSV → Consolidated Excel ─────────────────────
    def process_cash_ledger(self, input_paths):
        """Process Electronic Cash Ledger CSVs from GST portal."""
        try:
            e = _get_engine()
            all_data = []
            errors = []

            for p in input_paths:
                try:
                    data = e.parse_ecashl_csv(p)
                    all_data.append(data)
                except Exception as ex:
                    errors.append(f"{Path(p).name}: {ex}")

            if not all_data:
                return {'success': False, 'error': 'No valid Cash Ledger files. ' + '; '.join(errors)}

            out_name = f"CashLedger_Consolidated_{_ts()}.xlsx"
            out_path = os.path.join(self.output_folder, out_name)
            e.write_consolidated_ecashl(all_data, out_path)

            return {
                'success': True,
                'filename': out_name,
                'files_processed': len(all_data),
                'message': f'✅ {len(all_data)} Cash Ledger file(s) processed.'
            }

        except Exception as ex:
            return {'success': False, 'error': str(ex)}

    # ── PR vs 2B Reconciliation (uses GSTR-2B + Purchase Register) ──────────
    def reconcile_gstr2b(self, gstr2b_paths, pr_paths):
        """
        Reconcile Purchase Register vs GSTR-2B.
        For now redirects to GSTR-2B consolidation (full reco needs PR format).
        """
        try:
            all_paths = gstr2b_paths + pr_paths
            return self.process_gstr2b(all_paths)
        except Exception as ex:
            return {'success': False, 'error': str(ex)}

    # ── Proxy for old interface (called by existing app.py routes) ────────────
    def process_gstr1_single(self, input_path):
        return self.process_gstr1([input_path])

    def process_gstr2b_single(self, input_path):
        return self.process_gstr2b([input_path])

    def process_gstr3b_single(self, input_path):
        return self.process_gstr3b([input_path])

    def process_gstr9_single(self, input_path):
        return self.process_gstr9([input_path])

    def pr_vs_2b_reco(self, pr_path, gstr2b_path):
        return self.reconcile_gstr2b([gstr2b_path], [pr_path])
