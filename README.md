# Email Retrieval and Analysis

This Python script retrieves emails received today and saves relevant information to an ODS file. It can be useful for tasks like analyzing job-related emails.

## Prerequisites

- Python 3
- `openpyxl` library for working with Excel files
- An email account with IMAP access enabled

## Setup

1. **Clone the repository:**

    ```bash
    git clone https://github.com/your-username/email-retrieval.git
    ```

2. **Install dependencies:**

    ```bash
    pip install openpyxl
    ```

3. **Modify the `latest_email.py` file with your email address and password:**

    ```python
    # Input your email address and password
    email_address = "your_email@gmail.com"
    password = "your_password"
    ```

## Usage

Run the script:

```bash
python latest_email.py
