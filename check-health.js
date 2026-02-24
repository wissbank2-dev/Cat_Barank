require('dotenv').config();
const apiKey = process.env.GEMINI_API_KEY;

async function checkKumaHealth() {
    console.log('--- KUMA Heart Rate Check ---');
    console.log('API Key:', apiKey ? 'FOUND' : 'MISSING');

    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ contents: [{ parts: [{ text: "ping" }] }] })
        });
        const data = await response.json();

        if (data.error) {
            console.log('Pulse: WEAK 😿');
            console.log('Error Message:', data.error.message);
            console.log('Status Code:', data.error.status);
            console.log('Code:', data.error.code);
        } else {
            console.log('Pulse: STRONG! ✅');
            console.log('Response:', data.candidates[0].content.parts[0].text);
        }
    } catch (err) {
        console.error('Connection broken:', err.message);
    }
}

checkKumaHealth();
