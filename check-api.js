require('dotenv').config();
const apiKey = process.env.GEMINI_API_KEY;

async function checkModels() {
    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;
        const response = await fetch(url);
        const data = await response.json();

        if (data.error) {
            console.log('API Error:', JSON.stringify(data.error, null, 2));
        } else {
            console.log('Available Models:');
            data.models.forEach(m => console.log(`- ${m.name}`));
        }
    } catch (err) {
        console.error('Fetch error:', err.message);
    }
}

checkModels();
