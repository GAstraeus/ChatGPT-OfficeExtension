# ChatGPT-OfficeExtension

## Project Overview

This project integrates OpenAI's GPT-3 into Microsoft Office applications, transforming how users interact with digital content by leveraging the power of artificial intelligence. The GPT-3 model enhances productivity and creativity in applications like Word, Outlook, and PowerPoint through intuitive and dynamic text generation. This README provides an overview of the project, development insights, and instructions for setup and running the Office Add-On.

## Background

The GPT-3 integration project began as a vision to harness the AI's capabilities for enhancing Microsoft Office products. Our journey involved developing a Minimum Viable Product (MVP), confronting the challenges of add-on development, and exploring innovative solutions to integrate advanced AI technology into everyday productivity tools. The insights gained from developing, testing, and iterating on this add-on have paved the way for a more interactive and intelligent digital content creation experience.

## Implementation

The MVP allows users to access GPT-3 through an add-in pane in Microsoft Word, directly communicating with OpenAI's ChatGPT to generate and insert text based on user prompts. The project also navigated through the complexities of Manifest.xml in add-on development, emphasizing the importance of careful planning and consideration in the design and deployment of technology solutions.

## Getting Started

To get started with the GPT-3 Microsoft Office Add-On, follow these setup and running instructions:

1. **Clone the Repository:**
`git clone [repository-url]`


2. **Install Dependencies:**
Ensure you have the necessary dependencies installed, such as Node.js for running the local server and any other dependencies listed in the project's `package.json` you can use `npm install`

3. **Local Development Setup:**
- Start by setting up your local development environment to test and modify the add-on.
- If developing locally, ensure that the Manifest.xml file points to `localhost`. This is crucial as the add-on needs to fetch the right files from your development server.

4. **Running the Add-On:**
- To run the add-on locally, start your server with `npm run start`.
- This will automatically side load your add on into Microsoft Word
- Once the add-on is loaded, you should be able to interact with GPT-3 directly within your Microsoft Office application.

5. **Manifest File Adjustment:**
- The Manifest.xml file contains essential configuration information for your add-on. When developing locally, you must update the source location in the Manifest.xml file to point to your local server (usually `localhost` with a specific port).
- Remember, any changes in hosting or significant properties in the manifest file will require updating and redeploying the add-on.

## Contributing

We welcome contributions and suggestions to improve the GPT-3 Microsoft Office Add-On. Feel free to fork the repository, make your improvements, and submit a pull request with a clear explanation of your changes or enhancements.

## Blog Post

To read the blog post about this project see: https://sites.google.com/view/georgeaboudiwan/projects/hellotinker

## Useful Links
* https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/word-quickstart?tabs=yeomangenerator
* https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/word-tutorial