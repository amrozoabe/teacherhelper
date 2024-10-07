# Teacherhelper

Last updated: Mon Oct  7 09:21:33 UTC 2024

This project was successfully built.

﻿## TEACHERHELPER

## Privacy Policy

Effective Date: 11/09/2024

1. Introduction
Welcome to Teacherhelper ("we," "our," or "us"). We are committed to protecting your privacy and handling your data in an open and transparent manner. This privacy policy explains how we collect, use, share, and protect your personal information when you use our Excel add-in, Teacherhelper.

2. Information We Collect
2.1 Information You Provide
Teacherhelper does not collect or store any data from your Excel sheets, including student data or teacher's data. You retain full control over this information. 

2.2 Automatically Collected Information
We may automatically collect certain information about your use of Teacherhelper, including:

Log data (e.g., IP address, browser type, pages visited)

Device information (e.g., device type, operating system)

Usage data (e.g., features used, frequency of use)

3. How We Use Your Information
We use the information we collect for the following purposes:

To provide and maintain Teacherhelper's functionality

To improve and optimise Teacherhelper

To respond to your requests or inquiries

To comply with legal obligations

4. Third-Party Services and APIs
4.1 Anthropic Claude API
Teacherhelper allows you to use your own Anthropic Claude API for processing prompts and generating email drafts. Please note:

We do not have access to your Anthropic API key or the data processed through it.

You are responsible for the security of your API key.

The use of the Anthropic Claude API is subject to Anthropic's own privacy policy and terms of service. We recommend reviewing these documents at the relevant Anthropic webpage.

4.2 SendGrid API
Teacherhelper integrates with SendGrid for email sending functionality. Please note:

You will be using your own SendGrid API for sending emails.

We do not have access to your SendGrid API key or the emails sent through it.

You are responsible for the security of your API key and compliance with SendGrid's policies.

The use of the SendGrid API is subject to SendGrid's privacy policy and terms of service. We recommend reviewing these documents at the relevant SendGrid webpage.

5. Data Processing and Storage
Teacherhelper does not store any of your Excel data, student information, or email content. All processing occurs locally on your device or through your personal API integrations. 

6. Data Sharing and Disclosure
We do not sell, rent, or share your personal information with third parties except in the following circumstances:

With your consent

To comply with legal obligations

To protect our rights, privacy, safety, or property

7. Data Retention
We retain automatically collected data only for as long as necessary to fulfill the purposes outlined in this privacy policy, unless a longer retention period is required by law. 

8. Your Rights and Choices
Depending on your location, you may have certain rights regarding your personal information, including:

The right to access your data

The right to correct inaccurate data

The right to delete your data

The right to restrict or object to processing

The right to data portability

To exercise these rights, please contact us using the information provided in the "Contact Us" section.

9. Security
We implement appropriate technical and organizational measures to protect the limited information we automatically collect. However, please note that the security of your Excel data, API keys, and email content is your responsibility.

10. Changes to This Privacy Policy
We may update this privacy policy from time to time. We will notify you of any changes by posting the new privacy policy on this page and updating the "Effective Date" at the top.

11. Children's Privacy
Teacherhelper is not intended for use by children under the age of 13. We do not knowingly collect personal information from children under 13. If you believe we have collected personal information from a child under 13, please contact us immediately.

12. International Data Transfers
Your information may be transferred to and processed in countries other than your own. These countries may have data protection laws that are different from those in your country. By using Teacherhelper, you consent to the transfer of your information to these countries.

13. Contact Us
If you have any questions, concerns, or requests regarding this privacy policy or our data practices, please contact us at:

amrozoabe@teacherhelper.onmicrosoft.com

14. Governing Law and Jurisdiction
This privacy policy shall be governed by and construed in accordance with the laws of New South Wales, Australia. Any disputes arising under or in connection with this privacy policy shall be subject to the exclusive jurisdiction of the courts of New South Wales, Australia.

## How to Use Teacherhelper
1. teacherhelper is designed to streamline bulk communications from educators to students.
2. Select who should receive the email: all students or specific ones based on the available attributes. These attributes are represented as the column titles. You would need to select a specific value that appears within a column to filter based on it.
3. Choose the column containing email addresses. The first column containg emails is automatically selected. If more than one column contains email addresses, make sure to select the one you want to communicate with. Selecting a column that does not contain emails would result in an error. Rows with no valid email addresses will get skipped.
4. Write a prompt for the email content. Be direct.
5. Click "Generate Email Draft" to create a draft using Anthropic Claude AI. If you are not happy with the output, press the "Generate Email Draft" button again and repeat until you are satisfied.
6. Review and edit the generated subject and body. Make sure to maintain the column titles in the format {{column_title}} so they get personalised.
7. Set up your email signature. It will automatically be added below the email body. Once saved, it will continue to be automatically added until a new signature is added and saved.
8. You have the option of sending an email copy to yourself. You can write an email in the specified textarea and it will receive a copy of whatever emails goes out. This sender email can also be saved so it automatically appears when the add-in is loaded again.
9. Click "Send all emails" to send personalised emails to selected students. Please note that the students would receive an email from the following address: "no-reply.teacherhelper@outlook.com". If they decide to reply anyway, you would not receive their reply.
10. Please also note that you can only send up to 100 emails per day.
11. Advanced Settings allow users to have more control over the generated email draft. Further details are provided in the Advanced Settings section.

## Advanced settings instructions:
1. The controls above are described below:
2. The user can experiment with configuring the prompts that go to Claude AI.
3. The user can write a word, few words, or a sentence in each of the controls text area.
4. Institution, refers to the type of educational institution teacherhelper Excel Add-in is being used in. Default value is 'university'
5. Persona, refers to who should the email be assumed to be sent by. Default value is 'professor'
6. Audience refers to who would these emails be sent to, Default value is 'students'
7. Tone, refers to the writing style and tone that should be used to write this email. e.g. professional, casual, serious, etc.. Default value is 'professional'
8. When deciding the maximum response length, the term "token" is roughly equivalent to a word or subword. For example, 'hello' is one token, while 'unbelievable' might be broken into multiple tokens like 'un', 'believe', and 'able'. The default value is 300
9. "Predictability vs Creativity" is refering 'temperature' of the model. Low temperature (closer to 0) produces more predictable and conservative outputs. High temperature (closer to 1) increases randomness and creativity in the output. The default value is 0.7
