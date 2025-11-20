import os
import sys
import time
import importlib
from unittest import mock

# Ensure project path (insert repo root, regardless of execution cwd)
repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, repo_root)

# import run_gui App class
import run_gui

# create temp files
sample_csv = os.path.abspath('gsc_keyword_report_sample.csv')
if not os.path.exists(sample_csv):
    print('Sample CSV not found:', sample_csv)
    sys.exit(2)

# instantiate App without mainloop, then call load function
app = run_gui.App()
# avoid opening persistent watchers during this test
try:
    app._watch_stop = True
except Exception:
    pass

# set variables similar to UI
app.kws_var.set('gsc_keyword_report_sample.csv')
app.start_var.set('2025-10-01')
app.end_var.set('2025-10-31')
app.property_var.set('https://example.com')
app.outbase_var.set('gsc_keyword_report_test')
app.format_var.set('CSV')

# load csv into table
print('Loading CSV into table...')
app.load_csv_into_table(sample_csv)
print('Loaded rows:', len(app.current_rows), 'columns:', app.current_columns)

# patch filedialog for save path for CSV
save_csv_path = os.path.abspath('test_export.csv')
with mock.patch('tkinter.filedialog.asksaveasfilename', return_value=save_csv_path):
    print('Exporting CSV to', save_csv_path)
    app.export_csv()
    time.sleep(0.5)  # allow file write
    exists = os.path.exists(save_csv_path)
    print('CSV exported exists:', exists)
    print('Log tags after CSV export:', app.log.tag_names())

# patch filedialog for save path for row export and xlsx
save_row_path = os.path.abspath('test_row_export.csv')
with mock.patch('tkinter.filedialog.asksaveasfilename', return_value=save_row_path):
    if app.current_rows:
        print('Exporting first row to', save_row_path)
        app.export_row(app.current_rows[0])
        time.sleep(0.5)
        exists_row = os.path.exists(save_row_path)
        print('Row export exists:', exists_row)
        print('Log tags after row export:', app.log.tag_names())

# for xlsx export
save_xlsx_path = os.path.abspath('test_export.xlsx')
with mock.patch('tkinter.filedialog.asksaveasfilename', return_value=save_xlsx_path):
    # set to excel
    app.format_var.set('Excel (.xlsx)')
    # call export; since code will ask for pandas presence, ensure pandas available
    print('Exporting XLSX to', save_xlsx_path)
    app.export_csv()
    time.sleep(1)
    print('XLSX exported exists:', os.path.exists(save_xlsx_path))
    print('Log tags after XLSX export:', app.log.tag_names())

def test_on_run():
    print('\n--- Testing on_run ---')
    # Create dummy files for this test
    test_kws_path = os.path.abspath('test_on_run_kws.csv')
    test_sa_path = os.path.abspath('test_on_run_sa.json')
    with open(test_kws_path, 'w') as f:
        f.write('keyword1\nkeyword2')
    with open(test_sa_path, 'w') as f:
        f.write('{}')

    # Set UI variables for on_run test
    app.property_var.set('https://on-run-test.com')
    app.start_var.set('2025-11-01')
    app.end_var.set('2025-11-30')
    app.kws_var.set(test_kws_path)
    app.sa_var.set(test_sa_path)
    app.outbase_var.set('on_run_test_output')
    app.format_var.set('CSV')

    # Mock the backend script and messagebox
    with mock.patch('run_gui.messagebox') as mock_msgbox:
        # We need to mock the main function from the gsc_keyword_report module
        with mock.patch('gsc_keyword_report.main') as mock_gsc_main:
            print('Calling on_run...')
            app.on_run()
            # Wait for the thread to execute
            time.sleep(2)

            # Check that the backend script was called
            assert mock_gsc_main.called, "gsc_keyword_report.main was not called."
            
            # The arguments are passed via sys.argv, so we inspect that
            # The worker in on_run modifies sys.argv
            modified_argv = sys.argv
            print(f"sys.argv after on_run: {modified_argv}")

            # Check for expected arguments
            assert '--property' in modified_argv
            assert 'https://on-run-test.com' in modified_argv
            assert '--keywords' in modified_argv
            assert test_kws_path in modified_argv
            assert '--start-date' in modified_argv
            assert '2025-11-01' in modified_argv
            assert '--end-date' in modified_argv
            assert '2025-11-30' in modified_argv
            assert '--service-account' in modified_argv
            assert test_sa_path in modified_argv
            
            expected_output_filename = app.get_export_filename('.csv')
            assert '--output' in modified_argv
            assert expected_output_filename in modified_argv
            
            print('on_run test passed.')

    # Clean up dummy files
    os.remove(test_kws_path)
    os.remove(test_sa_path)

# Run the new test
test_on_run()

print('TEST COMPLETE')
# close app
try:
    app.destroy()
except Exception:
    pass
