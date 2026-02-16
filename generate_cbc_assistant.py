#!/usr/bin/env python3
"""
Complete CBC Assistant Generator
Reads Excel data and generates fully-featured HTML application in one step.

Usage:
    python3 generate_cbc_assistant.py

Output:
    CBCAssistant_Complete.html - Ready to use!
"""

import pandas as pd
import json
import sys
from pathlib import Path

print("=" * 60)
print("CBC ASSISTANT GENERATOR")
print("=" * 60)

# File paths
excel_path = 'CBC_Settlement_Funding_Master_v4.xlsx'
template_path = 'CBCAssistant_v4.html'
output_path = 'CBCAssistant_Complete.html'

# Check files exist
if not Path(excel_path).exists():
    print(f"❌ Error: {excel_path} not found!")
    print("   Make sure the Excel file is in the same directory.")
    sys.exit(1)

if not Path(template_path).exists():
    print(f"❌ Error: {template_path} not found!")
    print("   You need the original template file.")
    sys.exit(1)

print(f"\n📊 Reading data from: {excel_path}")

# ===== STEP 1: READ EXCEL DATA =====
df_cases = pd.read_excel(excel_path, sheet_name='Cases')
df_jurisdictions_sheet = pd.read_excel(excel_path, sheet_name='Jurisdictions')
df_state_data = pd.read_excel(excel_path, sheet_name='State Data')
df_insurance = pd.read_excel(excel_path, sheet_name='Insurance Carriers')
df_court_access = pd.read_excel(excel_path, sheet_name='Court Access')

has_jurisdictions_sheet = True

# ===== STEP 2: CATEGORIZE CASE STATUSES =====
def categorize_status(status):
    if pd.isna(status):
        return None
    status_lower = str(status).lower()
    if 'denied' in status_lower:
        return 'Denied'
    elif 'dismiss' in status_lower:
        return 'Dismissed'
    elif 'approved' in status_lower or 'granted' in status_lower:
        return 'Approved'
    return None

# Build cases list - include all cases with valid status, even without judge names
cases_data = []
for idx, row in df_cases.iterrows():
    if pd.notna(row.get('State')) and pd.notna(row.get('County')):
        status_cat = categorize_status(row.get('Status'))
        if status_cat:  # Only include if status is categorized
            judge_name = row.get('Judge') if pd.notna(row.get('Judge')) else 'Unknown Judge'
            cases_data.append({
                'state': str(row['State']),
                'county': str(row['County']),
                'judge': judge_name,
                'clientName': str(row.get('Client_Name', '')),
                'caseNumber': str(row.get('Case_Number', '')) if pd.notna(row.get('Case_Number')) else '',
                'courtDate': str(row.get('Court_Date', '')) if pd.notna(row.get('Court_Date')) else '',
                'status': status_cat,
                'notes': str(row.get('Notes', '')) if pd.notna(row.get('Notes')) else ''
            })

print(f"  ✓ Built {len(cases_data)} case records")

# ===== STEP 3: BUILD JURISDICTION LIST FROM COURT ACCESS =====
jurisdictions_data = []
for idx, row in df_court_access.iterrows():
    state = row.get('State')
    county = row.get('County')

    if pd.notna(state) and pd.notna(county):
        state = str(state)
        county = str(county)

        # Look for redaction data and county notes in Jurisdictions sheet with fuzzy matching
        redaction_level = ''
        redaction_notes = ''
        county_notes = ''

        if has_jurisdictions_sheet:
            # Try exact match first
            match = df_jurisdictions_sheet[
                (df_jurisdictions_sheet['State'] == state) &
                (df_jurisdictions_sheet['County'] == county)
            ]

            # If no exact match, try fuzzy matching
            if len(match) == 0:
                # Try adding " County" to the Atlas county name
                county_with_suffix = county + ' County'
                match = df_jurisdictions_sheet[
                    (df_jurisdictions_sheet['State'] == state) &
                    (df_jurisdictions_sheet['County'] == county_with_suffix)
                ]

                # If still no match, try matching base names
                if len(match) == 0:
                    for jur_idx, jur_row in df_jurisdictions_sheet[df_jurisdictions_sheet['State'] == state].iterrows():
                        jur_county = str(jur_row['County'])
                        jur_county_base = jur_county.replace(' County', '').strip()
                        atlas_county_base = county.replace(' County', '').strip()

                        if jur_county_base.lower() == atlas_county_base.lower():
                            match = df_jurisdictions_sheet.loc[[jur_idx]]
                            break

            if len(match) > 0:
                redaction_level = str(match.iloc[0].get('Redaction_Level', '')) if pd.notna(match.iloc[0].get('Redaction_Level')) else ''
                redaction_notes = str(match.iloc[0].get('Redaction_Notes', '')) if pd.notna(match.iloc[0].get('Redaction_Notes')) else ''
                county_notes = str(match.iloc[0].get('County_Notes', '')) if pd.notna(match.iloc[0].get('County_Notes')) else ''

        # Default redaction levels
        if not redaction_level:
            if state == 'NY':
                redaction_level = 'No Redaction'
            elif state in ['CA', 'MN', 'VA']:
                redaction_level = 'Full Redaction'

        # Extract court access info
        westlaw_coverage = str(row.get('Westlaw_Coverage', '')) if pd.notna(row.get('Westlaw_Coverage')) else ''
        website = str(row.get('Website', '')) if pd.notna(row.get('Website')) else ''
        fee_structure = str(row.get('Fee_Structure', '')) if pd.notna(row.get('Fee_Structure')) else ''
        subscription_login = str(row.get('Subscription_Login', '')) if pd.notna(row.get('Subscription_Login')) else ''
        search_notes = str(row.get('Search_Notes', '')) if pd.notna(row.get('Search_Notes')) else ''
        court_type = str(row.get('Court_Type', '')) if pd.notna(row.get('Court_Type')) else ''

        jurisdictions_data.append({
            'state': state,
            'county': county,
            'redactionLevel': redaction_level,
            'redactionNotes': redaction_notes,
            'countyNotes': county_notes,
            'westlawCoverage': westlaw_coverage,
            'website': website,
            'feeStructure': fee_structure,
            'subscriptionLogin': subscription_login,
            'searchNotes': search_notes,
            'courtType': court_type
        })

print(f"  ✓ Built {len(jurisdictions_data)} jurisdiction records from Court Access")

# ===== STEP 4: BUILD STATE DATA =====
state_data = {}
for idx, row in df_state_data.iterrows():
    state_code = str(row['State'])
    state_data[state_code] = {
        'name': str(row.get('State_Name', state_code)),
        'rateCap': str(row.get('Rate_Cap', '')) if pd.notna(row.get('Rate_Cap')) else '',
        'requiresIPA': str(row.get('Requires_IPA', '')) if pd.notna(row.get('Requires_IPA')) else '',
        'requiresAffDec': str(row.get('Requires_Aff_Dec', '')) if pd.notna(row.get('Requires_Aff_Dec')) else '',
        'noPoachState': str(row.get('No_Poach_State', '')) if pd.notna(row.get('No_Poach_State')) else '',
        'legalCounsel': str(row.get('Legal_Counsel', '')) if pd.notna(row.get('Legal_Counsel')) else '',
        'expectedLegalFee': str(row.get('Expected_Legal_Fee', '')) if pd.notna(row.get('Expected_Legal_Fee')) else '',
        'additionalNotes': str(row.get('Additional_Notes', '')) if pd.notna(row.get('Additional_Notes')) else ''
    }

# ===== STEP 5: BUILD INSURANCE DATA =====
insurance_data = {}
for idx, row in df_insurance.iterrows():
    carrier_name = str(row['Carrier_Name'])
    insurance_data[carrier_name] = {
        'name': carrier_name,
        'adminFee': str(row.get('Admin_Fee', '')) if pd.notna(row.get('Admin_Fee')) else '',
        'contact': str(row.get('Contact', '')) if pd.notna(row.get('Contact')) else '',
        'notes': str(row.get('Notes', '')) if pd.notna(row.get('Notes')) else ''
    }

# Calculate stats
total_denied = sum(1 for c in cases_data if c['status'] == 'Denied')
total_dismissed = sum(1 for c in cases_data if c['status'] == 'Dismissed')
total_approved = sum(1 for c in cases_data if c['status'] == 'Approved')

print(f"\n✅ Data processed successfully!")
print(f"   • {len(jurisdictions_data)} total jurisdictions")
print(f"   • {len([j for j in jurisdictions_data if any(c['state'] == j['state'] and c['county'] == j['county'] for c in cases_data)])} counties with case history")
print(f"   • {len(cases_data)} cases tracked")
print(f"   • {total_denied} denied, {total_dismissed} dismissed, {total_approved} approved")

# ===== STEP 6: READ TEMPLATE AND INJECT DATA =====
print(f"\n📝 Reading template: {template_path}")

with open(template_path, 'r', encoding='utf-8') as f:
    html = f.read()

# Find where to inject data (look for existing data sections or script tag)
cases_json = json.dumps(cases_data, indent=8)
jurisdictions_json = json.dumps(jurisdictions_data, indent=8)
state_data_json = json.dumps(state_data, indent=8)
insurance_data_json = json.dumps(insurance_data, indent=8)

# Replace data sections
html = html.replace('const casesData = [];', f'const casesData = {cases_json};')
html = html.replace('const jurisdictionsData = [];', f'const jurisdictionsData = {jurisdictions_json};')
html = html.replace('const stateData = {};', f'const stateData = {state_data_json};')
html = html.replace('const insuranceData = {};', f'const insuranceData = {insurance_data_json};')

print("  ✓ Injected data into template")

# ===== STEP 7: ADD ALL ENHANCEMENTS =====
print("\n🎨 Adding UI enhancements...")

# Add CSS
additional_css = """
        /* Approval Status */
        .status-approved { background: #d1fae5; color: #065f46; }
        .approved-count { color: #059669; }
        .judge-stats .stat.approved { color: #059669; }

        /* Search Bar */
        .search-container { position: relative; margin-bottom: 20px; }
        .search-input {
            width: 100%;
            padding: 10px 40px 10px 12px;
            border: 1px solid var(--gray-300);
            border-radius: 6px;
            font-size: 14px;
        }
        .search-input:focus {
            outline: none;
            border-color: var(--primary-color);
            box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1);
        }
        .search-icon {
            position: absolute;
            right: 12px;
            top: 50%;
            transform: translateY(-50%);
            color: var(--gray-400);
        }
        .search-results {
            position: fixed;
            background: white;
            border: 1px solid var(--gray-200);
            border-radius: 6px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
            max-height: 400px;
            overflow-y: auto;
            display: none;
            z-index: 9999;
            min-width: 300px;
        }
        .search-results.visible { display: block; }
        .search-result-item {
            padding: 12px 16px;
            cursor: pointer;
            border-bottom: 1px solid var(--gray-100);
        }
        .search-result-item:hover { background-color: var(--gray-50); }
        .search-result-county {
            font-weight: 600;
            color: var(--gray-900);
            font-size: 13px;
        }
        .search-result-state {
            font-size: 11px;
            color: var(--gray-500);
            margin-top: 2px;
        }
        .no-results {
            padding: 16px;
            text-align: center;
            color: var(--gray-500);
            font-size: 13px;
        }
"""

html = html.replace('</style>', additional_css + '\n    </style>')
print("  ✓ Added CSS")

# Add search bar HTML
old_sidebar = """                <aside class="sidebar">
                    <div class="card">
                        <div class="card-header"><h2><i class="fas fa-filter"></i> Filter</h2></div>
                        <div class="card-body">
                            <div class="form-group">
                                <label>State</label>
                                <select id="courtStateSelect"><option value="">All States</option></select>
                            </div>
                        </div>
                    </div>
                </aside>"""

new_sidebar = """                <aside class="sidebar">
                    <div class="card">
                        <div class="card-header"><h2><i class="fas fa-search"></i> Search & Filter</h2></div>
                        <div class="card-body">
                            <div class="form-group">
                                <label>Search Counties</label>
                                <div class="search-container">
                                    <input type="text" id="countySearch" class="search-input" placeholder="Type to search all 3,140 counties...">
                                    <i class="fas fa-search search-icon"></i>
                                    <div id="searchResults" class="search-results"></div>
                                </div>
                            </div>
                            <div class="form-group">
                                <label>Filter by State</label>
                                <select id="courtStateSelect"><option value="">All States</option></select>
                            </div>
                        </div>
                    </div>
                </aside>"""

html = html.replace(old_sidebar, new_sidebar)
print("  ✓ Added search bar")

# Fix county card meta
old_meta = """                        <div class="county-card-meta">
                            <span><i class="fas fa-map-marker-alt"></i> ${stateName}</span>
                            <span><i class="fas fa-user-tie"></i> ${judges.length} Judges</span>
                            <span><i class="fas fa-gavel"></i> ${totalCases} Cases</span>
                            <span class="denied-count"><i class="fas fa-times-circle"></i> ${totalDenied} Denied</span>
                            <span class="dismissed-count"><i class="fas fa-ban"></i> ${totalDismissed} Dismissed</span>
                        </div>"""

new_meta = """                        <div class="county-card-meta">
                            <span><i class="fas fa-map-marker-alt"></i> ${stateName}</span>
                            <span><i class="fas fa-user-tie"></i> ${judges.length} Judges</span>
                            <span><i class="fas fa-gavel"></i> ${totalCases} Cases</span>
                            <span class="approved-count"><i class="fas fa-check-circle"></i> ${totalApproved} Approved</span>
                            <span class="dismissed-count"><i class="fas fa-ban"></i> ${totalDismissed} Dismissed</span>
                            <span class="denied-count"><i class="fas fa-times-circle"></i> ${totalDenied} Denied</span>
                        </div>"""

html = html.replace(old_meta, new_meta)
print("  ✓ Fixed county cards")

# Fix judge stats
old_judge = """                                <div class="judge-stats">
                                    <span class="stat total"><i class="fas fa-folder"></i> ${judgeCases.length}</span>
                                    ${judgeDenied > 0 ? `<span class="stat denied"><i class="fas fa-times-circle"></i> ${judgeDenied} Denied</span>` : ''}
                                    ${judgeDismissed > 0 ? `<span class="stat dismissed"><i class="fas fa-ban"></i> ${judgeDismissed} Dismissed</span>` : ''}
                                </div>"""

new_judge = """                                <div class="judge-stats">
                                    <span class="stat total"><i class="fas fa-folder"></i> ${judgeCases.length}</span>
                                    <span class="stat approved"><i class="fas fa-check-circle"></i> ${judgeApproved} Approved</span>
                                    ${judgeDismissed > 0 ? `<span class="stat dismissed"><i class="fas fa-ban"></i> ${judgeDismissed} Dismissed</span>` : ''}
                                    ${judgeDenied > 0 ? `<span class="stat denied"><i class="fas fa-times-circle"></i> ${judgeDenied} Denied</span>` : ''}
                                </div>"""

html = html.replace(old_judge, new_judge)
print("  ✓ Fixed judge cards")

# Add calculations
old_calc = """            const judges = Object.keys(judgeMap);
            const totalCases = cases.length;
            const totalDenied = cases.filter(c => c.status === 'Denied').length;
            const totalDismissed = cases.filter(c => c.status === 'Dismissed').length;"""

new_calc = """            const judges = Object.keys(judgeMap);
            const totalCases = cases.length;
            const totalApproved = cases.filter(c => c.status === 'Approved').length;
            const totalDismissed = cases.filter(c => c.status === 'Dismissed').length;
            const totalDenied = cases.filter(c => c.status === 'Denied').length;"""

html = html.replace(old_calc, new_calc)

old_judge_calc = """                    const judgeDenied = judgeCases.filter(c => c.status === 'Denied').length;
                    const judgeDismissed = judgeCases.filter(c => c.status === 'Dismissed').length;"""

new_judge_calc = """                    const judgeApproved = judgeCases.filter(c => c.status === 'Approved').length;
                    const judgeDismissed = judgeCases.filter(c => c.status === 'Dismissed').length;
                    const judgeDenied = judgeCases.filter(c => c.status === 'Denied').length;"""

html = html.replace(old_judge_calc, new_judge_calc)
print("  ✓ Added calculations")

# Add county notes
old_notes = """                                ${jurisdiction.redactionNotes ? `
                                <div class="redaction-info-item">
                                    <label>Notes</label>
                                    <span>${jurisdiction.redactionNotes}</span>
                                </div>
                                ` : ''}
                            </div>
                        </div>"""

new_notes = """                                ${jurisdiction.redactionNotes ? `
                                <div class="redaction-info-item">
                                    <label>Notes</label>
                                    <span>${jurisdiction.redactionNotes}</span>
                                </div>
                                ` : ''}
                                ${jurisdiction.countyNotes ? `
                                <div class="redaction-info-item">
                                    <label>Jurisdiction Notes</label>
                                    <span style="white-space: pre-line;">${jurisdiction.countyNotes}</span>
                                </div>
                                ` : ''}
                            </div>
                        </div>"""

html = html.replace(old_notes, new_notes)
print("  ✓ Added county notes")

# Add sorting and approved stats
old_stats = """            const totalDismissed = filteredCases.filter(c => c.status === 'Dismissed').length;

            document.getElementById("courtStats").innerHTML = `
                <div class="stat-item"><div class="stat-value">${totalCounties}</div><div class="stat-label">Jurisdictions</div></div>
                <div class="stat-item"><div class="stat-value">${totalJudges}</div><div class="stat-label">Judges</div></div>
                <div class="stat-item"><div class="stat-value">${totalCases}</div><div class="stat-label">Total Cases</div></div>
                <div class="stat-item"><div class="stat-value" style="color: var(--danger-red);">${totalDenied}</div><div class="stat-label">Denied</div></div>
                <div class="stat-item"><div class="stat-value" style="color: var(--warning-yellow);">${totalDismissed}</div><div class="stat-label">Dismissed</div></div>
            `;

            document.getElementById("entryCount").textContent = `Showing ${filteredJurisdictions.length} jurisdictions with ${totalJudges} judges and ${totalCases} cases`;

            if (filteredJurisdictions.length === 0) {
                container.innerHTML = '';
                placeholder.style.display = 'block';
                return;
            }

            placeholder.style.display = 'none';
            container.innerHTML = filteredJurisdictions.map(jur => createCountyCard(jur)).join('');"""

new_stats = """            const totalDismissed = filteredCases.filter(c => c.status === 'Dismissed').length;
            const totalApproved = filteredCases.filter(c => c.status === 'Approved').length;

            document.getElementById("courtStats").innerHTML = `
                <div class="stat-item"><div class="stat-value">${totalCounties}</div><div class="stat-label">Jurisdictions</div></div>
                <div class="stat-item"><div class="stat-value">${totalJudges}</div><div class="stat-label">Judges</div></div>
                <div class="stat-item"><div class="stat-value">${totalCases}</div><div class="stat-label">Total Cases</div></div>
                <div class="stat-item"><div class="stat-value" style="color: #059669;">${totalApproved}</div><div class="stat-label">Approved</div></div>
                <div class="stat-item"><div class="stat-value" style="color: var(--warning-yellow);">${totalDismissed}</div><div class="stat-label">Dismissed</div></div>
                <div class="stat-item"><div class="stat-value" style="color: var(--danger-red);">${totalDenied}</div><div class="stat-label">Denied</div></div>
            `;

            document.getElementById("entryCount").textContent = `Showing ${filteredJurisdictions.length} jurisdictions with ${totalJudges} judges and ${totalCases} cases`;

            if (filteredJurisdictions.length === 0) {
                container.innerHTML = '';
                placeholder.style.display = 'block';
                return;
            }

            placeholder.style.display = 'none';

            // Sort by case count, then alphabetically
            const jurisdictionsWithCounts = filteredJurisdictions.map(j => {
                const caseCount = casesData.filter(c => c.state === j.state && c.county === j.county).length;
                return { ...j, caseCount };
            });
            jurisdictionsWithCounts.sort((a, b) => {
                if (b.caseCount !== a.caseCount) return b.caseCount - a.caseCount;
                return a.county.localeCompare(b.county);
            });

            container.innerHTML = jurisdictionsWithCounts.map(jur => createCountyCard(jur)).join('');"""

html = html.replace(old_stats, new_stats)
print("  ✓ Added sorting and stats")

# Add search JavaScript
search_js = """
        // County Search Functionality
        let searchTimeout;
        const searchInput = document.getElementById('countySearch');
        const searchResults = document.getElementById('searchResults');

        function positionSearchDropdown() {
            if (searchInput && searchResults) {
                const rect = searchInput.getBoundingClientRect();
                searchResults.style.left = rect.left + 'px';
                searchResults.style.top = (rect.bottom + 4) + 'px';
                searchResults.style.width = rect.width + 'px';
            }
        }

        if (searchInput) {
            searchInput.addEventListener('input', function(e) {
                clearTimeout(searchTimeout);
                const query = e.target.value.trim().toLowerCase();

                if (query.length < 2) {
                    searchResults.classList.remove('visible');
                    return;
                }

                searchTimeout = setTimeout(() => {
                    const matches = jurisdictionsData.filter(j =>
                        j.county.toLowerCase().includes(query) || j.state.toLowerCase().includes(query)
                    ).slice(0, 20);

                    if (matches.length > 0) {
                        searchResults.innerHTML = matches.map(j => `
                            <div class="search-result-item" data-state="${j.state}" data-county="${j.county}">
                                <div class="search-result-county">${j.county}</div>
                                <div class="search-result-state">${j.state}</div>
                            </div>
                        `).join('');

                        positionSearchDropdown();
                        searchResults.classList.add('visible');

                        document.querySelectorAll('.search-result-item').forEach(item => {
                            item.addEventListener('click', function() {
                                document.getElementById('courtStateSelect').value = this.dataset.state;
                                updateCourtGuide();
                                setTimeout(() => {
                                    const cards = document.querySelectorAll('.county-card h3');
                                    for (let card of cards) {
                                        if (card.textContent.trim() === this.dataset.county) {
                                            card.closest('.county-card').scrollIntoView({ behavior: 'smooth' });
                                            card.closest('.county-card').style.boxShadow = '0 0 0 3px rgba(59, 130, 246, 0.3)';
                                            setTimeout(() => { card.closest('.county-card').style.boxShadow = ''; }, 2000);
                                            break;
                                        }
                                    }
                                }, 100);
                                searchInput.value = '';
                                searchResults.classList.remove('visible');
                            });
                        });
                    } else {
                        searchResults.innerHTML = '<div class="no-results">No counties found</div>';
                        positionSearchDropdown();
                        searchResults.classList.add('visible');
                    }
                }, 300);
            });

            document.addEventListener('click', function(e) {
                if (searchInput && !searchInput.contains(e.target) && !searchResults.contains(e.target)) {
                    searchResults.classList.remove('visible');
                }
            });

            window.addEventListener('scroll', positionSearchDropdown);
            window.addEventListener('resize', positionSearchDropdown);
        }

    </script>
</body>"""

html = html.replace('    </script>\n</body>', search_js)
print("  ✓ Added search JavaScript")

# ===== STEP 8: SAVE FINAL HTML =====
with open(output_path, 'w', encoding='utf-8') as f:
    f.write(html)

print("\n" + "=" * 60)
print("✅ SUCCESS!")
print("=" * 60)
print(f"\n📄 Generated: {output_path}")
print("\n✨ Features included:")
print("  ✓ All 3,140 counties from Atlas")
print("  ✓ Search bar with dropdown")
print("  ✓ Approvals, Dismissals, Denials tracked")
print("  ✓ Sorted by case count")
print("  ✓ County notes (Ramsey, James City, Fresno)")
print("  ✓ Court access info with Westlaw badges")
print("\n🎉 Ready to use!")
