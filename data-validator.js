
// data-validation.js - Data validation for Breaches Tracker
;(function() {
	'use strict';
	
	// Validate parsed data for integrity
	function validateData(data) {
		const errors = [];
		const warnings = [];
		
		if (!data || !Array.isArray(data) || data.length === 0) {
			errors.push('No data found in file');
			return { valid: false, errors, warnings };
		}
		
		const required = ['id', 'title', 'category', 'status', 'created', 'owner', 'agent',
		'TSLA breach', 'lead', 'content', 'language'];
		
		// Check first row has required columns
		const first = data[0] || {};
		const missing = required.filter(f => !(f in first));
		if (missing.length > 0) {
			errors.push('Missing required columns: ' + missing.join(', '));
		}
		
		// Row-level checks
		let emptyRows = 0;
		let malformedIds = 0;
		
		data.forEach((row, index) => {
			let hasData = false;
			Object.values(row).forEach(v => {
				if (v && String(v).trim()) hasData = true;
			});
			if (!hasData) emptyRows++;
			
			if (row.id && !String(row.id).trim()) malformedIds++;
		});
		
		if (emptyRows > 0) warnings.push(emptyRows + ' empty row(s) detected');
		if (malformedIds > 0) warnings.push(malformedIds + ' row(s) with missing ID');
		
		return { valid: errors.length === 0, errors, warnings };
	}
	
	// Display validation results
	function showValidation(result) {
		const container = document.getElementById('validation-results');
		if (!container) return;
		
		if (!result.valid) {
			container.innerHTML = '<p class="validation-error">' +
				result.errors.map(e => escapeHTML(e)).join('<br>') + '</p>';
			container.style.display = 'block';
			return false;
		}
		
		if (result.warnings.length > 0) {
			container.innerHTML = '<p class="validation-warning">' +
				result.warnings.join('<br>') + '</p>';
			container.style.display = 'block';
		} else {
			container.innerHTML = '<p class="validation-success">Validation passed!</p>';
			container.style.display = 'block';
		}
		return true;
	}
	
	function escapeHTML(str) {
		const d = document.createElement('div');
		d.textContent = String(str);
		return d.innerHTML;
	}
	
	window.BT = window.BT || {};
	window.BT.validateData = validateData;
	window.BT.showValidation = showValidation;
})();
