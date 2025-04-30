# Excel Uploader

A modern web application for uploading and processing multiple Excel files, built with Next.js and ready to deploy on Vercel.

## Features

- 🌟 Beautiful and responsive UI with Tailwind CSS
- 📁 Drag-and-drop file upload with React Dropzone
- 📊 Support for Excel files (.xls, .xlsx)
- 🚀 Upload up to 20 files simultaneously
- ⚡ Serverless API endpoint for file processing
- 🔒 Validation for file types and limits
- 📱 Mobile-friendly design

## Getting Started

### Prerequisites

- Node.js 18.x or later
- npm or yarn

### Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/excel-uploader.git
cd excel-uploader
```

2. Install dependencies:
```bash
npm install
# or
yarn install
```

3. Run the development server:
```bash
npm run dev
# or
yarn dev
```

4. Open [http://localhost:3000](http://localhost:3000) in your browser to see the application.

## Deployment

The application is ready to be deployed on Vercel. Simply push your code to a Git repository and import it into Vercel.

[![Deploy with Vercel](https://vercel.com/button)](https://vercel.com/new/clone?repository-url=https%3A%2F%2Fgithub.com%2Fyourusername%2Fexcel-uploader)

## Tech Stack

- [Next.js](https://nextjs.org/) - React framework for server-rendered applications
- [React](https://reactjs.org/) - JavaScript library for building user interfaces
- [Tailwind CSS](https://tailwindcss.com/) - Utility-first CSS framework
- [React Dropzone](https://react-dropzone.js.org/) - React hook for file uploads
- [TypeScript](https://www.typescriptlang.org/) - Typed JavaScript

## Project Structure

```
excel-uploader/
├── public/             # Static assets
├── src/
│   ├── app/            # App router pages and layouts
│   │   ├── api/        # API routes
│   │   │   └── upload/ # File upload endpoint
│   │   ├── layout.tsx  # Root layout
│   │   └── page.tsx    # Home page
│   ├── components/     # Reusable components
│   │   └── FileUpload.tsx  # File upload component
│   └── ...
├── package.json        # Project dependencies
└── ...
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- [Vercel](https://vercel.com) for the amazing deployment platform
- [Next.js](https://nextjs.org) team for the fantastic framework
