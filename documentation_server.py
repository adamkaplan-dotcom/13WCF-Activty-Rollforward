"""
13WCF Activity Rollforward - Documentation Dashboard Server
===========================================================
A Flask web server that provides an interactive documentation dashboard.

Usage:
    python3 documentation_server.py

Then open your browser to: http://localhost:5002
"""

from flask import Flask, render_template_string
import os
import webbrowser
from threading import Timer

app = Flask(__name__)

# HTML template for the dashboard
DASHBOARD_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>13WCF Activity Rollforward - Documentation Dashboard</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            min-height: 100vh;
            padding: 20px;
        }

        .dashboard-container {
            max-width: 1400px;
            margin: 0 auto;
        }

        .header {
            background: white;
            padding: 30px;
            border-radius: 15px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
            margin-bottom: 30px;
        }

        .header h1 {
            color: #1e3c72;
            font-size: 2.5em;
            margin-bottom: 10px;
        }

        .header p {
            color: #666;
            font-size: 1.1em;
        }

        .header .version {
            display: inline-block;
            background: #1e3c72;
            color: white;
            padding: 5px 15px;
            border-radius: 20px;
            font-size: 0.9em;
            margin-top: 10px;
        }

        .main-grid {
            display: grid;
            grid-template-columns: 300px 1fr;
            gap: 30px;
        }

        .sidebar {
            background: white;
            padding: 25px;
            border-radius: 15px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
            height: fit-content;
            position: sticky;
            top: 20px;
        }

        .sidebar h3 {
            color: #1e3c72;
            margin-bottom: 20px;
            font-size: 1.3em;
        }

        .nav-item {
            padding: 12px 15px;
            margin-bottom: 8px;
            border-radius: 8px;
            cursor: pointer;
            transition: all 0.3s ease;
            color: #333;
            font-weight: 500;
        }

        .nav-item:hover {
            background: #f0f0f0;
            transform: translateX(5px);
        }

        .nav-item.active {
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            color: white;
        }

        .content-area {
            background: white;
            padding: 40px;
            border-radius: 15px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
            min-height: 600px;
        }

        .section {
            display: none;
        }

        .section.active {
            display: block;
            animation: fadeIn 0.5s;
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .section h2 {
            color: #1e3c72;
            font-size: 2em;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 3px solid #1e3c72;
        }

        .section h3 {
            color: #2a5298;
            font-size: 1.5em;
            margin-top: 30px;
            margin-bottom: 15px;
        }

        .section h4 {
            color: #333;
            font-size: 1.2em;
            margin-top: 20px;
            margin-bottom: 10px;
        }

        .section p {
            line-height: 1.8;
            color: #555;
            margin-bottom: 15px;
        }

        .section ul, .section ol {
            margin-left: 30px;
            margin-bottom: 20px;
            line-height: 1.8;
            color: #555;
        }

        .section li {
            margin-bottom: 8px;
        }

        .code-block {
            background: #f8f9fa;
            border-left: 4px solid #1e3c72;
            padding: 20px;
            border-radius: 8px;
            margin: 20px 0;
            font-family: 'Monaco', 'Courier New', monospace;
            overflow-x: auto;
            white-space: pre;
            font-size: 0.9em;
            line-height: 1.6;
        }

        .inline-code {
            background: #f0f0f0;
            padding: 3px 8px;
            border-radius: 4px;
            font-family: 'Monaco', 'Courier New', monospace;
            font-size: 0.9em;
            color: #c7254e;
        }

        .info-box {
            background: #d1ecf1;
            border-left: 4px solid #0c5460;
            padding: 20px;
            margin: 20px 0;
            border-radius: 8px;
        }

        .warning-box {
            background: #fff3cd;
            border-left: 4px solid #856404;
            padding: 20px;
            margin: 20px 0;
            border-radius: 8px;
        }

        .success-box {
            background: #d4edda;
            border-left: 4px solid #155724;
            padding: 20px;
            margin: 20px 0;
            border-radius: 8px;
        }

        .file-card {
            background: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            margin: 15px 0;
            border: 2px solid #e0e0e0;
        }

        .file-card h4 {
            color: #1e3c72;
            margin-top: 0;
        }

        .step-box {
            background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
            padding: 20px;
            border-radius: 10px;
            margin: 15px 0;
            border-left: 5px solid #1e3c72;
        }

        .step-box h4 {
            color: #1e3c72;
            margin-top: 0;
            margin-bottom: 10px;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }

        table th, table td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #e0e0e0;
        }

        table th {
            background: #f8f9fa;
            color: #1e3c72;
            font-weight: 600;
        }

        .badge {
            display: inline-block;
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 0.85em;
            font-weight: 600;
            margin-right: 10px;
        }

        .badge-primary {
            background: #1e3c72;
            color: white;
        }

        .badge-required {
            background: #dc3545;
            color: white;
        }

        .badge-optional {
            background: #6c757d;
            color: white;
        }

        .badge-success {
            background: #28a745;
            color: white;
        }
    </style>
</head>
<body>
    <div class="dashboard-container">
        <div class="header">
            <h1>⚙️ 13WCF Activity Rollforward App</h1>
            <p>Interactive documentation for the Balance Updater & Activity Rollforward automation tool</p>
            <span class="version">v2.6.1</span>
        </div>

        <div class="main-grid">
            <div class="sidebar">
                <h3>Navigation</h3>
                <div class="nav-item active" onclick="showSection('overview')">📖 Overview</div>
                <div class="nav-item" onclick="showSection('quickstart')">🚀 Quick Start</div>
                <div class="nav-item" onclick="showSection('files')">📁 Required Files</div>
                <div class="nav-item" onclick="showSection('processing')">⚙️ Processing Steps</div>
                <div class="nav-item" onclick="showSection('interface')">🖥️ Web Interface</div>
                <div class="nav-item" onclick="showSection('architecture')">🏗️ Architecture</div>
                <div class="nav-item" onclick="showSection('modules')">📦 Modules</div>
                <div class="nav-item" onclick="showSection('troubleshooting')">🔧 Troubleshooting</div>
                <div class="nav-item" onclick="showSection('technical')">🔬 Technical Details</div>
            </div>

            <div class="content-area">
                <!-- Overview Section -->
                <div id="overview" class="section active">
                    <h2>Project Overview</h2>
                    <p>The <strong>13WCF Activity Rollforward App</strong> is a Python Flask web application that automates the processing of weekly balance data and activity rollforward updates for the Bullish 13-Week Cash Flow model.</p>

                    <div class="info-box">
                        <strong>💡 Key Benefits:</strong>
                        <ul style="margin-top: 10px; margin-left: 20px;">
                            <li>Eliminates manual copy-paste operations</li>
                            <li>Automatically adjusts Excel formulas</li>
                            <li>Processes multiple worksheets in one operation</li>
                            <li>Handles large files (up to 1GB total)</li>
                            <li>Updates FvA data tabs from 13WCF files</li>
                        </ul>
                    </div>

                    <h3>What It Does</h3>
                    <p>The application performs <strong>12 distinct processing steps</strong>:</p>

                    <div class="step-box">
                        <h4>Steps 1-3: Balance Updates</h4>
                        <p>Updates three worksheets with current week balances from the Weekly Balances file:</p>
                        <ul>
                            <li><strong>Beginning Balances</strong> - Starting balance data</li>
                            <li><strong>Cockpit</strong> - Summary and control data</li>
                            <li><strong>Rollforward Check</strong> - Validation and reconciliation</li>
                        </ul>
                    </div>

                    <div class="step-box">
                        <h4>Steps 4-11: Stacked Activity Updates</h4>
                        <p>Processes activity data from the Activity Aggregator file and updates the Activity Rollforward file with stacked activity calculations.</p>
                    </div>

                    <div class="step-box">
                        <h4>Step 12: FvA Data Tabs</h4>
                        <p>Updates FvA (Fair Value Adjustment) data tabs for 1-Week, 4-Week, and 13-Week periods using data from 13WCF Consolidated files (optional).</p>
                    </div>

                    <h3>Technology Stack</h3>
                    <ul>
                        <li><strong>Python 3.9+</strong> - Core runtime environment</li>
                        <li><strong>Flask 3.0.0</strong> - Web framework for the interface</li>
                        <li><strong>openpyxl 3.1.2</strong> - Excel file manipulation</li>
                        <li><strong>pandas</strong> - Data processing and analysis</li>
                        <li><strong>numpy</strong> - Numerical operations</li>
                    </ul>
                </div>

                <!-- Quick Start Section -->
                <div id="quickstart" class="section">
                    <h2>Quick Start Guide</h2>

                    <h3>Starting the Application</h3>
                    <div class="code-block">Double-click: START_PYTHON_APP.command</div>
                    <p>The launcher script will automatically:</p>
                    <ol>
                        <li>Check for Python 3 installation</li>
                        <li>Create/activate a virtual environment</li>
                        <li>Install required dependencies</li>
                        <li>Start the Flask web server on port 5000</li>
                    </ol>

                    <div class="success-box">
                        <strong>✅ Server Started!</strong><br>
                        Open your browser and go to: <strong>http://localhost:5000</strong>
                    </div>

                    <h3>Basic Workflow</h3>
                    <ol>
                        <li><strong>Upload Required Files</strong> - Drag & drop or click to browse
                            <ul>
                                <li>Weekly Balances (required)</li>
                                <li>Activity Rollforward (required)</li>
                                <li>BTH Investments (required)</li>
                                <li>Activity Aggregator (required)</li>
                            </ul>
                        </li>
                        <li><strong>Upload Optional FvA Files</strong> (if needed)
                            <ul>
                                <li>1-Week 13WCF file</li>
                                <li>4-Week 13WCF file</li>
                                <li>13-Week 13WCF file</li>
                            </ul>
                        </li>
                        <li><strong>Click "Process Files"</strong> - Processing takes 10-30 seconds</li>
                        <li><strong>Review Processing Log</strong> - Check for any errors or warnings</li>
                        <li><strong>Download Updated File</strong> - File includes timestamp in filename</li>
                    </ol>

                    <h3>Stopping the Server</h3>
                    <p>Press <strong>CTRL+C</strong> in the terminal window to stop the server.</p>
                </div>

                <!-- Files Section -->
                <div id="files" class="section">
                    <h2>Required Files</h2>

                    <div class="file-card">
                        <h4>1. Weekly Balances <span class="badge badge-required">Required</span></h4>
                        <p><strong>Typical filename:</strong> <span class="inline-code">Weekly_Balances_YYYY-MM-DD.xlsx</span></p>
                        <p><strong>Required tab:</strong> "USDx Balances"</p>
                        <p><strong>Data used:</strong> Column K (Current Week balances) starting from row 13</p>
                        <p><strong>Purpose:</strong> Source file containing the latest weekly balance data for all accounts.</p>
                    </div>

                    <div class="file-card">
                        <h4>2. Activity Rollforward <span class="badge badge-required">Required</span></h4>
                        <p><strong>Typical filename:</strong> <span class="inline-code">Activity_Rollforward.xlsx</span></p>
                        <p><strong>Required tabs:</strong> "Beginning Balances", "Cockpit", "Rollforward Check"</p>
                        <p><strong>Purpose:</strong> Master file that gets updated with new balance and activity data. This is the main output file.</p>
                        <table>
                            <thead>
                                <tr>
                                    <th>Tab</th>
                                    <th>Update Method</th>
                                    <th>Key Row</th>
                                </tr>
                            </thead>
                            <tbody>
                                <tr>
                                    <td>Beginning Balances</td>
                                    <td>Adds new column after last date in row 14</td>
                                    <td>14</td>
                                </tr>
                                <tr>
                                    <td>Cockpit</td>
                                    <td>Adds new column after last date in row 12</td>
                                    <td>12</td>
                                </tr>
                                <tr>
                                    <td>Rollforward Check</td>
                                    <td>Adds new column after last date in row 5</td>
                                    <td>5</td>
                                </tr>
                            </tbody>
                        </table>
                    </div>

                    <div class="file-card">
                        <h4>3. BTH Investments <span class="badge badge-required">Required</span></h4>
                        <p><strong>Typical filename:</strong> <span class="inline-code">BTH_Investments.xlsx</span></p>
                        <p><strong>Purpose:</strong> Contains Buy-To-Hold investment data used in balance calculations.</p>
                    </div>

                    <div class="file-card">
                        <h4>4. Activity Aggregator <span class="badge badge-required">Required</span></h4>
                        <p><strong>Typical filename:</strong> <span class="inline-code">Activity_Aggregator.xlsx</span></p>
                        <p><strong>Purpose:</strong> Source file for stacked activity data that gets processed and added to the Activity Rollforward file.</p>
                        <p><strong>Note:</strong> Large file (typically 50-60MB).</p>
                    </div>

                    <h3>Optional FvA Files</h3>

                    <div class="file-card">
                        <h4>5. 1-Week 13WCF <span class="badge badge-optional">Optional</span></h4>
                        <p><strong>Required tab:</strong> "13WCF - Consol"</p>
                        <p><strong>Purpose:</strong> Updates 1-Week FvA data tab in Activity Rollforward file.</p>
                    </div>

                    <div class="file-card">
                        <h4>6. 4-Week 13WCF <span class="badge badge-optional">Optional</span></h4>
                        <p><strong>Required tab:</strong> "13WCF - Consol"</p>
                        <p><strong>Purpose:</strong> Updates 4-Week FvA data tab in Activity Rollforward file.</p>
                    </div>

                    <div class="file-card">
                        <h4>7. 13-Week 13WCF <span class="badge badge-optional">Optional</span></h4>
                        <p><strong>Required tab:</strong> "13WCF - Consol"</p>
                        <p><strong>Purpose:</strong> Updates 13-Week FvA data tab in Activity Rollforward file.</p>
                    </div>

                    <div class="warning-box">
                        <strong>⚠️ File Size Limit:</strong> Total upload size for all files is limited to 1GB. Individual files can be quite large, but ensure the combined size doesn't exceed this limit.
                    </div>
                </div>

                <!-- Processing Section -->
                <div id="processing" class="section">
                    <h2>Processing Steps Explained</h2>

                    <h3>Steps 1-3: Balance Updates</h3>

                    <div class="step-box">
                        <h4>Step 1: Beginning Balances Tab</h4>
                        <p><strong>What happens:</strong></p>
                        <ol>
                            <li>Finds last date in row 14 of Beginning Balances tab</li>
                            <li>Creates new column after the last date</li>
                            <li>Copies date formula from row 10 of previous column (adjusted for new column)</li>
                            <li>Pastes Current Week balances from Weekly Balances file (Column K)</li>
                            <li>Matches balances by reference number in Column B</li>
                            <li>Copies formulas from previous column for rows 12-20 and 148-end</li>
                            <li>Copies cell formatting from previous column</li>
                        </ol>
                    </div>

                    <div class="step-box">
                        <h4>Step 2: Cockpit Tab</h4>
                        <p><strong>What happens:</strong></p>
                        <ol>
                            <li>Finds last date in row 12 of Cockpit tab</li>
                            <li>Creates new column after the last date</li>
                            <li>Copies date formula from row 10 of previous column</li>
                            <li>Pastes Current Week balances matched by Column B</li>
                            <li>Copies formulas from previous column for rows 14-22 and 148-end</li>
                        </ol>
                    </div>

                    <div class="step-box">
                        <h4>Step 3: Rollforward Check Tab</h4>
                        <p><strong>What happens:</strong></p>
                        <ol>
                            <li>Finds last date in row 5 of Rollforward Check tab</li>
                            <li>Creates new column after the last date</li>
                            <li>Copies date formula from row 3 of previous column</li>
                            <li>Pastes Current Week balances matched by Column B</li>
                            <li>Copies formulas from previous column for rows 7-15 and 138-end</li>
                        </ol>
                    </div>

                    <h3>Steps 4-11: Stacked Activity Processing</h3>

                    <div class="info-box">
                        <strong>ℹ️ Module:</strong> <span class="inline-code">stacked_activity_updater.py</span>
                        <p style="margin-top: 10px;">Processes activity data from the Activity Aggregator and updates the Activity Rollforward file with stacked activity calculations. This is a complex multi-step process handled by a dedicated module.</p>
                    </div>

                    <h3>Step 12: FvA Data Tab Updates</h3>

                    <div class="info-box">
                        <strong>ℹ️ Module:</strong> <span class="inline-code">fva_data_updater.py</span>
                        <p style="margin-top: 10px;">If FvA files are provided, reads "13WCF - Consol" sheet from each and updates corresponding FvA Data tabs (1-Week, 4-Week, 13-Week) in the Activity Rollforward file.</p>
                    </div>

                    <p><strong>Only runs if:</strong> At least one of the three FvA files (1-Week, 4-Week, or 13-Week) is uploaded.</p>
                </div>

                <!-- Interface Section -->
                <div id="interface" class="section">
                    <h2>Web Interface Guide</h2>

                    <h3>Upload Area</h3>
                    <p>The interface displays 7 upload zones in a responsive grid layout:</p>

                    <h4>Required Files (Red border)</h4>
                    <ul>
                        <li>Weekly Balances</li>
                        <li>Activity Rollforward</li>
                        <li>BTH Investments</li>
                        <li>Activity Aggregator</li>
                    </ul>

                    <h4>Optional Files (Gray border)</h4>
                    <ul>
                        <li>1-Week 13WCF</li>
                        <li>4-Week 13WCF</li>
                        <li>13-Week 13WCF</li>
                    </ul>

                    <h3>Drag & Drop Features</h3>
                    <ul>
                        <li><strong>Visual Feedback:</strong> Drop zones highlight when you drag files over them</li>
                        <li><strong>File Preview:</strong> Shows filename after successful upload</li>
                        <li><strong>Change File:</strong> Click the zone again to replace the file</li>
                        <li><strong>Multiple Files:</strong> Each zone accepts one file; upload all before processing</li>
                    </ul>

                    <h3>Process Button</h3>
                    <div class="code-block">Process Files</div>
                    <p>Only enabled when all 4 required files are uploaded. Optional FvA files can be included or omitted.</p>

                    <h3>Processing Log</h3>
                    <p>During processing, a detailed log appears showing:</p>
                    <ul>
                        <li>Each step being executed (1-12)</li>
                        <li>Row counts and data statistics</li>
                        <li>Success confirmations with ✓ checkmarks</li>
                        <li>Any warnings or errors with ❌ markers</li>
                        <li>Total processing time</li>
                    </ul>

                    <h3>Download Section</h3>
                    <p>After successful processing:</p>
                    <ul>
                        <li>Download link appears automatically</li>
                        <li>Filename format: <span class="inline-code">Activity_Rollforward_Updated_YYYYMMDD_HHMMSS.xlsx</span></li>
                        <li>File is saved in <span class="inline-code">outputs/</span> directory</li>
                        <li>Previous output files remain accessible</li>
                    </ul>
                </div>

                <!-- Architecture Section -->
                <div id="architecture" class="section">
                    <h2>Architecture & File Structure</h2>

                    <h3>Directory Structure</h3>
                    <div class="code-block">BalanceUpdater/
├── 13WCF-Activity Rollforward App v2.6.1.py  # Main Flask app
├── stacked_activity_updater.py                # Stacked activity module
├── fva_data_updater.py                        # FvA update module
├── requirements.txt                           # Python dependencies
├── START_PYTHON_APP.command                   # Launch script
├── VERSION                                    # Version history
├── CLAUDE.md                                  # Technical documentation
├── README.md                                  # User guide
├── templates/
│   └── index.html                            # Web interface
├── uploads/                                  # Temp file storage (auto-deleted)
├── outputs/                                  # Processed files
└── python_venv/                              # Virtual environment</div>

                    <h3>Request Flow</h3>
                    <ol>
                        <li><strong>User uploads files</strong> → Saved to <span class="inline-code">uploads/</span> with secure filenames</li>
                        <li><strong>User clicks Process</strong> → Flask POST request to <span class="inline-code">/upload</span></li>
                        <li><strong>Main app coordinates</strong> → Calls processing functions in sequence</li>
                        <li><strong>Modules execute</strong> → Balance updater → Stacked activity → FvA updater</li>
                        <li><strong>Output saved</strong> → Timestamped file in <span class="inline-code">outputs/</span></li>
                        <li><strong>Temp files deleted</strong> → <span class="inline-code">uploads/</span> cleaned automatically</li>
                        <li><strong>Download link returned</strong> → User downloads processed file</li>
                    </ol>

                    <h3>API Endpoints</h3>
                    <table>
                        <thead>
                            <tr>
                                <th>Endpoint</th>
                                <th>Method</th>
                                <th>Purpose</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td><span class="inline-code">/</span></td>
                                <td>GET</td>
                                <td>Serve main upload interface</td>
                            </tr>
                            <tr>
                                <td><span class="inline-code">/upload</span></td>
                                <td>POST</td>
                                <td>Process uploaded files</td>
                            </tr>
                            <tr>
                                <td><span class="inline-code">/download/&lt;filename&gt;</span></td>
                                <td>GET</td>
                                <td>Download processed file</td>
                            </tr>
                            <tr>
                                <td><span class="inline-code">/outputs</span></td>
                                <td>GET</td>
                                <td>List all output files</td>
                            </tr>
                        </tbody>
                    </table>
                </div>

                <!-- Modules Section -->
                <div id="modules" class="section">
                    <h2>Processing Modules</h2>

                    <h3>Main Application</h3>
                    <p><strong>File:</strong> <span class="inline-code">13WCF-Activity Rollforward App v2.6.1.py</span></p>
                    <p>Handles:</p>
                    <ul>
                        <li>Flask web server and routing</li>
                        <li>File upload handling and validation</li>
                        <li>Steps 1-3: Balance updates for all three tabs</li>
                        <li>Coordination of module execution</li>
                        <li>Output file generation and download</li>
                    </ul>

                    <h3>Stacked Activity Updater</h3>
                    <p><strong>File:</strong> <span class="inline-code">stacked_activity_updater.py</span></p>
                    <p>Handles Steps 4-11:</p>
                    <ul>
                        <li>Reads activity data from Activity Aggregator file</li>
                        <li>Processes and stacks activity calculations</li>
                        <li>Updates Activity Rollforward with stacked data</li>
                        <li>Complex formula adjustments and data transformations</li>
                    </ul>

                    <h3>FvA Data Updater</h3>
                    <p><strong>File:</strong> <span class="inline-code">fva_data_updater.py</span></p>
                    <p>Handles Step 12:</p>
                    <ul>
                        <li>Reads "13WCF - Consol" sheet from 13WCF files</li>
                        <li>Uses streaming XML processing for memory efficiency</li>
                        <li>Updates FvA Data tabs (1-Week, 4-Week, 13-Week)</li>
                        <li>Handles large files without memory issues</li>
                    </ul>
                </div>

                <!-- Troubleshooting Section -->
                <div id="troubleshooting" class="section">
                    <h2>Troubleshooting</h2>

                    <h3>Server Won't Start</h3>
                    <p><strong>Error:</strong> "ModuleNotFoundError: No module named 'flask'"</p>
                    <p><strong>Solution:</strong></p>
                    <div class="code-block">pip3 install flask</div>
                    <p>Or run the START_PYTHON_APP.command script which installs dependencies automatically.</p>

                    <h3>Upload Errors</h3>

                    <h4>"413 Request Entity Too Large"</h4>
                    <p><strong>Cause:</strong> Total file size exceeds 1GB limit</p>
                    <p><strong>Solution:</strong> Check file sizes, ensure combined total is under 1GB</p>

                    <h4>"File type not allowed"</h4>
                    <p><strong>Cause:</strong> File is not .xlsx or .xls format</p>
                    <p><strong>Solution:</strong> Only upload Excel files (.xlsx or .xls)</p>

                    <h3>Processing Errors</h3>

                    <h4>"Sheet not found" errors</h4>
                    <p><strong>Cause:</strong> Required worksheet tab is missing or misnamed</p>
                    <p><strong>Solution:</strong> Verify file has required tabs:
                        <ul>
                            <li>Weekly Balances: "USDx Balances"</li>
                            <li>Activity Rollforward: "Beginning Balances", "Cockpit", "Rollforward Check"</li>
                            <li>13WCF files: "13WCF - Consol"</li>
                        </ul>
                    </p>

                    <h4>"No dates found in row X"</h4>
                    <p><strong>Cause:</strong> Expected date row is empty or formatted incorrectly</p>
                    <p><strong>Solution:</strong> Check that the date row in the Activity Rollforward file has at least one date value</p>

                    <h4>Formula errors (#REF!, #VALUE!)</h4>
                    <p><strong>Cause:</strong> Formula adjustment didn't work correctly</p>
                    <p><strong>Solution:</strong> Verify source formulas use relative references, not absolute row references</p>

                    <h3>Performance Issues</h3>

                    <h4>Processing takes very long (&gt; 2 minutes)</h4>
                    <p><strong>Possible causes:</strong></p>
                    <ul>
                        <li>Very large files (Activity Aggregator &gt; 100MB)</li>
                        <li>Thousands of rows being processed</li>
                        <li>Low system memory</li>
                    </ul>
                    <p><strong>Solutions:</strong></p>
                    <ul>
                        <li>Close other applications to free up memory</li>
                        <li>Check processing log to see which step is slow</li>
                        <li>Wait for completion - the app handles large files but may take time</li>
                    </ul>

                    <h3>Output File Issues</h3>

                    <h4>Can't open output file in Excel</h4>
                    <p><strong>Cause:</strong> File may be corrupted or incomplete</p>
                    <p><strong>Solution:</strong> Check processing log for errors, try processing again</p>

                    <h4>Missing data in output</h4>
                    <p><strong>Cause:</strong> Reference matching issue or empty source data</p>
                    <p><strong>Solution:</strong> Verify Column B reference numbers match between files</p>
                </div>

                <!-- Technical Details Section -->
                <div id="technical" class="section">
                    <h2>Technical Details</h2>

                    <h3>openpyxl Configuration</h3>
                    <div class="code-block">Reading source files:
- data_only=True     # Gets calculated values, not formulas
- read_only=True     # Memory efficiency for large files

Writing target file:
- Default mode       # Preserves formulas and formatting</div>

                    <h3>Key Row Numbers</h3>
                    <p><strong>Note:</strong> In code, rows are 0-indexed (subtract 1 from Excel row number)</p>

                    <table>
                        <thead>
                            <tr>
                                <th>Worksheet</th>
                                <th>Purpose</th>
                                <th>Excel Row</th>
                                <th>Code Index</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>Beginning Balances</td>
                                <td>Date row</td>
                                <td>14</td>
                                <td>13</td>
                            </tr>
                            <tr>
                                <td>Beginning Balances</td>
                                <td>Formula header</td>
                                <td>10</td>
                                <td>9</td>
                            </tr>
                            <tr>
                                <td>Cockpit</td>
                                <td>Date row</td>
                                <td>12</td>
                                <td>11</td>
                            </tr>
                            <tr>
                                <td>Cockpit</td>
                                <td>Formula header</td>
                                <td>10</td>
                                <td>9</td>
                            </tr>
                            <tr>
                                <td>Rollforward Check</td>
                                <td>Date row</td>
                                <td>5</td>
                                <td>4</td>
                            </tr>
                            <tr>
                                <td>Rollforward Check</td>
                                <td>Formula header</td>
                                <td>3</td>
                                <td>2</td>
                            </tr>
                            <tr>
                                <td>USDx Balances (source)</td>
                                <td>Header row</td>
                                <td>10</td>
                                <td>9</td>
                            </tr>
                            <tr>
                                <td>USDx Balances (source)</td>
                                <td>Data starts</td>
                                <td>13</td>
                                <td>12</td>
                            </tr>
                        </tbody>
                    </table>

                    <h3>Formula Copying Ranges</h3>

                    <h4>Beginning Balances</h4>
                    <ul>
                        <li>Rows 12-20 (code: 11-19)</li>
                        <li>Rows 148-end (code: 147-end)</li>
                    </ul>

                    <h4>Cockpit</h4>
                    <ul>
                        <li>Rows 14-22 (code: 13-21)</li>
                        <li>Rows 148-end (code: 147-end)</li>
                    </ul>

                    <h4>Rollforward Check</h4>
                    <ul>
                        <li>Rows 7-15 (code: 6-14)</li>
                        <li>Rows 138-end (code: 137-end)</li>
                    </ul>

                    <h3>Reference Matching</h3>
                    <p>Balance data is matched between files using reference numbers in <strong>Column B</strong>.</p>
                    <p>The app skips section headers by checking for keywords:</p>
                    <div class="code-block">Skip rows containing (case-insensitive):
- 'bullish'
- 'coindesk'
- 'fiat'
- 'balance'
- 'total'</div>

                    <h3>Cell Formatting</h3>
                    <p>When copying formulas, the app also copies cell formatting:</p>
                    <ul>
                        <li>Font (style, size, color, bold, italic)</li>
                        <li>Border (style, color)</li>
                        <li>Fill (background color)</li>
                        <li>Number format (currency, percentage, etc.)</li>
                        <li>Alignment (horizontal, vertical)</li>
                        <li>Protection (locked/unlocked)</li>
                    </ul>

                    <h3>Security</h3>
                    <ul>
                        <li><strong>Local only:</strong> Server runs on localhost, no external access</li>
                        <li><strong>Secure filenames:</strong> Uses <span class="inline-code">secure_filename()</span> for all uploads</li>
                        <li><strong>Auto-cleanup:</strong> Uploaded files deleted immediately after processing</li>
                        <li><strong>No authentication:</strong> Designed for local use only</li>
                    </ul>

                    <h3>Performance</h3>
                    <ul>
                        <li><strong>Typical processing time:</strong> 10-30 seconds for standard files</li>
                        <li><strong>Large file handling:</strong> Can process files up to 100MB+</li>
                        <li><strong>Memory usage:</strong> Scales with file size and row count</li>
                        <li><strong>Concurrency:</strong> Processes one request at a time</li>
                    </ul>
                </div>
            </div>
        </div>
    </div>

    <script>
        function showSection(sectionId) {
            // Hide all sections
            const sections = document.querySelectorAll('.section');
            sections.forEach(section => section.classList.remove('active'));

            // Remove active from all nav items
            const navItems = document.querySelectorAll('.nav-item');
            navItems.forEach(item => item.classList.remove('active'));

            // Show selected section
            document.getElementById(sectionId).classList.add('active');

            // Highlight selected nav item
            event.target.classList.add('active');

            // Scroll to top of content area
            document.querySelector('.content-area').scrollTop = 0;
        }
    </script>
</body>
</html>
"""


@app.route('/')
def dashboard():
    """Serve the documentation dashboard"""
    return render_template_string(DASHBOARD_HTML)


def open_browser():
    """Open the browser after a short delay"""
    webbrowser.open('http://localhost:5002')


if __name__ == '__main__':
    print("\n" + "="*60)
    print("13WCF Activity Rollforward - Documentation Dashboard")
    print("="*60)
    print("\nStarting server...")
    print("Dashboard will be available at: http://localhost:5002")
    print("\nPress Ctrl+C to stop the server")
    print("="*60 + "\n")

    # Open browser after 1 second delay
    Timer(1, open_browser).start()

    # Run Flask server
    app.run(host='localhost', port=5002, debug=False)
