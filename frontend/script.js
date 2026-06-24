const API_BASE = 'http://localhost:8000/api';

document.addEventListener('DOMContentLoaded', function() {
    let currentProposal = null;
    let currentTopic = '';
    
    const tabButtons = document.querySelectorAll('.tab-btn');
    
    tabButtons.forEach(function(btn) {
        btn.addEventListener('click', function() {
            const tabName = btn.dataset.tab;
            
            tabButtons.forEach(function(b) {
                b.classList.remove('active');
            });
            
            document.querySelectorAll('.tab-content').forEach(function(c) {
                c.classList.remove('active');
            });
            
            btn.classList.add('active');
            document.getElementById(tabName + '-tab').classList.add('active');
        });
    });
    
    function valueToString(value) {
        if (typeof value === 'string') {
            return value;
        } else if (typeof value === 'object') {
            return jsonToPlainText(value);
        }
        return String(value);
    }
    
    function jsonToPlainText(obj) {
        let parts = [];
        for (let key in obj) {
            if (typeof obj[key] === 'object') {
                parts.push(key + '\n' + jsonToPlainText(obj[key]));
            } else {
                parts.push(key + '\n' + obj[key]);
            }
        }
        return parts.join('\n\n');
    }
    
    async function downloadProposal(format) {
        try {
            console.log(`Downloading proposal as ${format}...`);
            const response = await fetch(`${API_BASE}/download-proposal`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    topic: currentTopic,
                    proposal: currentProposal,
                    format: format
                })
            });
            
            if (!response.ok) {
                const errorText = await response.text();
                throw new Error('Server error: ' + errorText);
            }
            
            const blob = await response.blob();
            console.log('Blob received, size:', blob.size, 'type:', blob.type);
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            const suffix = format === 'txt' ? '.txt' : (format === 'html' ? '.html' : '.docx');
            a.download = `${currentTopic.replace(/\s+/g, '_')}_Research_Proposal${suffix}`;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
            console.log('Download initiated');
        } catch (err) {
            console.error('Download error:', err);
            alert('Failed to download document: ' + err.message);
        }
    }
    
    document.getElementById('topic-form').addEventListener('submit', async function(e) {
        e.preventDefault();
        
        const dept = document.getElementById('department').value;
        const degree = document.getElementById('degree').value;
        const interest = document.getElementById('interest').value;
        const results = document.getElementById('topic-results');
        
        results.innerHTML = '<div class="loading">Generating topics...</div>';
        
        try {
            const response = await fetch(API_BASE + '/generate-topics', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    department: dept,
                    degree_level: degree,
                    research_interest: interest
                })
            });
            
            if (!response.ok) {
                const errorText = await response.text();
                throw new Error('Server error: ' + errorText);
            }
            
            const data = await response.json();
            console.log('API Response:', data);
            
            const topicsData = JSON.parse(data.topics);
            const topics = topicsData.topics;
            console.log('Parsed topics:', topics);
            
            let html = '';
            for (let i = 0; i < topics.length; i++) {
                const t = topics[i];
                html += '<div class="topic-card">';
                html += '<h3>' + (i + 1) + '. ' + valueToString(t.title || t.topic) + '</h3>';
                html += '<div class="scores">';
                html += '<span class="score novelty">Novelty: ' + (t.novelty_score || t.novelty) + '%</span>';
                html += '<span class="score feasibility">Feasibility: ' + (t.feasibility_score || t.feasibility) + '%</span>';
                html += '</div>';
                html += '<p class="gap"><strong>Research Gap:</strong> ' + valueToString(t.research_gap || t.gap) + '</p>';
                html += '</div>';
            }
            
            results.innerHTML = html;
        } catch (err) {
            console.error('Error:', err);
            results.innerHTML = '<p style="color: red;">Error: ' + err.message + '</p>';
        }
    });
    
    document.getElementById('proposal-form').addEventListener('submit', async function(e) {
        e.preventDefault();
        
        const topic = document.getElementById('topic').value;
        const dept = document.getElementById('proposal-department').value;
        const degree = document.getElementById('proposal-degree').value;
        const results = document.getElementById('proposal-results');
        
        currentTopic = topic;
        results.innerHTML = '<div class="loading">Generating proposal...</div>';
        
        try {
            const response = await fetch(API_BASE + '/generate-proposal', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    topic: topic,
                    department: dept,
                    degree_level: degree
                })
            });
            
            if (!response.ok) {
                const errorText = await response.text();
                throw new Error('Server error: ' + errorText);
            }
            
            const data = await response.json();
            console.log('API Response:', data);
            
            currentProposal = JSON.parse(data.proposal);
            console.log('Parsed proposal:', currentProposal);
            
            let html = `
                <div style="margin-bottom: 20px; display: flex; gap: 10px; flex-wrap: wrap;">
                    <button id="download-html" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; padding: 12px 24px; border-radius: 8px; font-size: 1rem; font-weight: 600; cursor: pointer;">🌐 Download as HTML (Recommended)</button>
                    <button id="download-txt" style="background: #555; color: white; border: none; padding: 12px 24px; border-radius: 8px; font-size: 1rem; font-weight: 600; cursor: pointer;">📄 Download as TXT</button>
                    <button id="download-docx" style="background: #555; color: white; border: none; padding: 12px 24px; border-radius: 8px; font-size: 1rem; font-weight: 600; cursor: pointer;">📝 Download as Word (DOCX)</button>
                </div>
            `;
            const sectionNames = {
                'background': '1.0 BACKGROUND OF THE STUDY',
                'problem_statement': '2.0 PROBLEM STATEMENT',
                'objectives': '3.0 OBJECTIVES OF THE STUDY',
                'research_questions': '4.0 RESEARCH QUESTIONS',
                'scope': '5.0 SCOPE OF THE STUDY',
                'significance': '6.0 SIGNIFICANCE OF THE STUDY',
                'literature_review': '7.0 LITERATURE REVIEW',
                'methodology': '8.0 METHODOLOGY',
                'expected_outcomes': '9.0 EXPECTED OUTCOMES'
            };
            
            for (const [key, value] of Object.entries(currentProposal)) {
                const displayName = sectionNames[key] || key.charAt(0).toUpperCase() + key.slice(1);
                html += '<div class="proposal-section">';
                html += '<h3>' + displayName + '</h3>';
                const contentStr = valueToString(value).replace(/\n/g, '<br>');
                html += '<p>' + contentStr + '</p>';
                html += '</div>';
            }
            
            results.innerHTML = html;
            
            document.getElementById('download-html').addEventListener('click', function() {
                downloadProposal('html');
            });
            document.getElementById('download-txt').addEventListener('click', function() {
                downloadProposal('txt');
            });
            document.getElementById('download-docx').addEventListener('click', function() {
                downloadProposal('docx');
            });
        } catch (err) {
            console.error('Error:', err);
            results.innerHTML = '<p style="color: red;">Error: ' + err.message + '</p>';
        }
    });
});
