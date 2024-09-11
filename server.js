const express = require('express');
const cors = require('cors');
const axios = require('axios');
const sgMail = require('@sendgrid/mail');

const app = express();
app.use(cors());
app.use(express.json());

// const ANTHROPIC_API_KEY = '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';
// const SENDGRID_API_KEY = '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';

app.post('/api/generate', async (req, res) => {
  console.log('Received request on /api/generate route');
  console.log('Request body:', JSON.stringify(req.body, null, 2));

  const { ANTHROPIC_API_KEY } = req.body;

  if (!ANTHROPIC_API_KEY) {
    return res.status(400).json({ error: 'Anthropic API key is required' });
  }

  try {
    console.log('Attempting to call Anthropic API...');
    const response = await axios.post('https://api.anthropic.com/v1/complete', {
      model: "claude-2",
      prompt: req.body.prompt,
      max_tokens_to_sample: req.body.max_tokens_to_sample || 300,
      temperature: req.body.temperature || 0.7,
    }, {
      headers: {
        'Content-Type': 'application/json',
        'X-API-Key': ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01'
      },
    });

    console.log('Anthropic API response:', JSON.stringify(response.data, null, 2));
    res.json(response.data);
  } catch (error) {
    console.error('Error calling Anthropic API:', error);
    if (error.response) {
      console.error('Anthropic API error response:', JSON.stringify(error.response.data, null, 2));
      console.error('Anthropic API error status:', error.response.status);
      console.error('Anthropic API error headers:', JSON.stringify(error.response.headers, null, 2));
    } else if (error.request) {
      console.error('No response received from Anthropic API:', error.request);
    } else {
      console.error('Error setting up the request:', error.message);
    }
    res.status(500).json({ 
      error: 'Error calling Anthropic API', 
      details: error.response ? error.response.data : error.message 
    });
  }
});

app.post('/api/send-emails', async (req, res) => {
  console.log('Received request to send emails');
  try {
    const { emails, SENDGRID_API_KEY, sendgridEmail } = req.body;
    console.log('Emails to send:', JSON.stringify(emails, null, 2));

    if (!emails || !Array.isArray(emails) || emails.length === 0) {
      console.log('Invalid or empty emails array');
      return res.status(400).json({ error: 'Invalid or empty emails array' });
    }

    if (!SENDGRID_API_KEY || !sendgridEmail) {
      return res.status(400).json({ error: 'SendGrid API key and email are required' });
    }

    sgMail.setApiKey(SENDGRID_API_KEY);

    const messages = emails.map(email => ({
      to: email.to,
      from: sendgridEmail,
      subject: email.subject,
      html: email.html,
    }));

    console.log('Prepared messages:', JSON.stringify(messages, null, 2));

    const sentEmails = [];
    const failedEmails = [];

    for (const message of messages) {
      try {
        console.log(`Attempting to send email to: ${message.to}`);
        const result = await sgMail.send(message);
        console.log(`SendGrid response for ${message.to}:`, JSON.stringify(result, null, 2));
        sentEmails.push(message.to);
        console.log(`Email sent successfully to ${message.to}`);
      } catch (error) {
        console.error(`Error sending email to ${message.to}:`, error);
        failedEmails.push({ email: message.to, error: error.message });
      }
    }

    console.log(`Successfully sent ${sentEmails.length} out of ${emails.length} emails`);
    console.log('Failed emails:', JSON.stringify(failedEmails, null, 2));
    
    res.json({ 
      message: "Emails sending process completed",
      sentEmails: sentEmails,
      failedEmails: failedEmails,
      totalEmails: emails.length
    });
  } catch (error) {
    console.error('Detailed error:', error);
    if (error.response) {
      console.error('SendGrid error response:', JSON.stringify(error.response.body, null, 2));
    }
    res.status(500).json({ 
      error: 'An error occurred while sending emails', 
      message: error.message,
      stack: error.stack,
      sendGridError: error.response ? error.response.body : null
    });
  }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));