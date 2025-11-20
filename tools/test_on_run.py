
import os
import sys
import time
import unittest
from unittest import mock
import tkinter as tk

# Ensure project path is set correctly
repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, repo_root)

import run_gui
import gsc_keyword_report

class TestRunGuiOnRun(unittest.TestCase):

    def setUp(self):
        """Set up the test environment."""
        self.app = run_gui.App()
        # Stop the file watcher during tests
        self.app._watch_stop = True
        
        # Create dummy files required for the test
        self.sample_kws_path = os.path.join(repo_root, 'test_keywords.csv')
        self.service_account_path = os.path.join(repo_root, 'test_sa.json')
        self.output_path = os.path.join(repo_root, 'test_output.csv')

        with open(self.sample_kws_path, 'w', encoding='utf-8') as f:
            f.write('test keyword 1\n')
            f.write('test keyword 2\n')
        
        with open(self.service_account_path, 'w', encoding='utf-8') as f:
            f.write('{"test": "account"}')

        # Clean up old output file if it exists
        if os.path.exists(self.output_path):
            os.remove(self.output_path)

    def tearDown(self):
        """Tear down the test environment."""
        # Destroy the app window
        self.app.destroy()
        
        # Remove dummy files
        if os.path.exists(self.sample_kws_path):
            os.remove(self.sample_kws_path)
        if os.path.exists(self.service_account_path):
            os.remove(self.service_account_path)
        if os.path.exists(self.output_path):
            os.remove(self.output_path)

    @mock.patch('run_gui.messagebox')
    @mock.patch('gsc_keyword_report.main')
    def test_on_run_call(self, mock_gsc_main, mock_messagebox):
        """Test if on_run correctly calls gsc_keyword_report.main with the right arguments."""
        
        # Set UI variables
        self.app.property_var.set('https://example.com')
        self.app.start_var.set('2025-01-01')
        self.app.end_var.set('2025-01-31')
        self.app.kws_var.set(self.sample_kws_path)
        self.app.sa_var.set(self.service_account_path)
        self.app.outbase_var.set('test_output_base')
        
        # The get_export_filename method will generate the output filename
        # We need to predict it to check the arguments
        expected_output_filename = self.app.get_export_filename('.csv')
        
        # We need to manually set the output var for the test
        # because on_run calculates it internally
        self.app.output_path_for_test = expected_output_filename

        # Call the on_run method
        self.app.on_run()

        # Allow the worker thread to run
        time.sleep(2) 

        # Check if gsc_keyword_report.main was called
        self.assertTrue(mock_gsc_main.called, "gsc_keyword_report.main was not called.")

        # Check the arguments passed to gsc_keyword_report.main
        # sys.argv is modified by the run_gui.py script
        call_args = sys.argv
        self.assertIn('--property', call_args)
        self.assertIn('https://example.com', call_args)
        self.assertIn('--keywords', call_args)
        self.assertIn(self.sample_kws_path, call_args)
        self.assertIn('--start-date', call_args)
        self.assertIn('2025-01-01', call_args)
        self.assertIn('--end-date', call_args)
        self.assertIn('2025-01-31', call_args)
        self.assertIn('--service-account', call_args)
        self.assertIn(self.service_account_path, call_args)
        self.assertIn('--output', call_args)
        self.assertIn(expected_output_filename, call_args)


if __name__ == '__main__':
    unittest.main()
