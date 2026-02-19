# 🤖 IAgent

A Rust-based AI agent designed for automated Excel spreadsheet manipulation, powered by the Deepseek API. This project explores the integration of LLMs with structured data processing in a systems programming environment.

## ✨ Technologies

- `Deepseek API`
- `Serde` (for serialization/deserialization)
- `Tokio` (for async runtime)
- `calamine` (Excel manipulation libraries)
- `rust_xlsxwriter` (Create and write Excel files)

## 🚀 Features

- **AI-Driven Commands**: Manipulate spreadsheets using natural language intent via Deepseek.
- **Excel Integration**: Read and write data directly to `.xlsx` files.
- **Async Performance**: Built with Rust's async ecosystem for efficient API handling.
- **Type Safety**: Leveraging Rust's compiler to ensure robust data transformations.

## 📍 The Process

I've always been curious about how AI can handle boring office tasks like spreadsheet management. While Python is the 'standard' for this, I wanted to see if I could build a faster, more robust version using Rust. The main challenge was mapping the AI's natural language output to specific Excel grid operations. I spent a lot of time fine-tuning the prompts to ensure the agent wouldn't break the cell formatting. It's still a work in progress, but it's a great proof of concept for automated data workflows.

## 🚦 Running the Project

1. Clone the repository:
   ```bash
   git clone https://github.com/lucascirille/IAgent.git
   cd IAgent
   ```
2. Set up your Api Key:
   ```bash
   export DEEPSEEK_API_KEY='your_api_key_here'
   
   "or set up in .env file"
   ```
3. Install dependences and run:
   `cargo run`

And that is all!, enjoy!.
