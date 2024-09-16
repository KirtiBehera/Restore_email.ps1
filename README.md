# Email Restore Automation Script

## Overview

This PowerShell script provides a streamlined, efficient way to restore emails from **user mailboxes**, **shared mailboxes**, or **any mailbox** in a Microsoft Exchange or Office 365 environment. It significantly reduces the effort required by administrators, enhances the organizational workflow, and marks a step forward toward automation and AI-driven solutions in IT management.

## Features

- **Restore Emails**: Easily restore emails from any mailbox (user, shared, etc.).
- **Flexible**: Allows the selection of different mailboxes for restoration.
- **Automation**: Simplifies the admin task, reducing manual intervention.
- **AI Integration**: Incorporates intelligent operations to automate tasks, improving workflow efficiency.
- **Error Handling**: Contains robust error-handling mechanisms for improved stability.
- **Logging**: Keeps logs of the restoration processes for auditing purposes.

## Benefits

- **Reduced Admin Effort**: Minimizes the manual effort required to restore emails.
- **Organization-Friendly**: Helps maintain a smooth organizational workflow.
- **Automation-First**: Embeds principles of automation, setting the stage for AI-driven tasks.
- **Scalable**: Suitable for small or large-scale environments.

## Requirements

- **PowerShell**: Ensure PowerShell 7.1 or later is installed.
- **Administrator Rights**: The script must be executed with administrator privileges.
- **Exchange Online PowerShell Module** or **Exchange Server Management Tools** must be installed.
- **Office 365 Account**: Admin credentials for accessing user mailboxes, if working with Office 365.
  
  **Note**: For on-premises environments, adjust settings accordingly to use the Exchange Management Shell.

## Setup and Installation

1. Clone this repository to your local machine:

   ```bash
   git clone https://https://github.com/KirtiBehera/Restore_email.ps1.git
   ```

2. Install the required PowerShell modules if they are not already installed:
   
   ```powershell
   Install-Module -Name ExchangeOnlineManagement
   ```

3. Make sure the script is placed in a secure location and has the necessary permissions.

4. Open PowerShell with administrator privileges.

5. Connect to Exchange Online or the required Exchange Server environment:

   ```powershell
   Connect-ExchangeOnline -UserPrincipalName <admin-email>
   ```

## Usage

1. **Running the Script**:

   Run the script by providing the necessary arguments for mailbox and restoration settings:

   ```powershell
   ./RestoreEmail.ps1 -Mailbox <mailbox@example.com> -StartDate <yyyy-mm-dd> -EndDate <yyyy-mm-dd>
   ```

2. **Arguments**:

   - `-Mailbox`: The email address of the mailbox to restore emails from.
   - `-StartDate`: The start date of the email range to restore.
   - `-EndDate`: The end date of the email range to restore.

   Example:

   ```powershell
   ./RestoreEmail.ps1 -Mailbox john.doe@example.com -StartDate 2024-01-01 -EndDate 2024-02-01
   ```

3. **Restore from a Shared Mailbox**:

   ```powershell
   ./RestoreEmail.ps1 -Mailbox shared@domain.com -StartDate 2023-07-01 -EndDate 2023-08-01
   ```

## Logging

Logs are automatically generated and saved in the `logs/` directory in the following format:

```
EmailRestore_<timestamp>.log
```

Each log file contains details of the restoration process, including:

- Mailbox address
- Date range
- Number of restored emails
- Any errors encountered

## Error Handling

The script includes error handling to:

- Retry operations if connectivity issues are detected.
- Catch and log exceptions without stopping the entire process.

## Contributing

Feel free to contribute to this project by submitting issues or pull requests. Contributions that improve the efficiency, performance, or security of the script are welcome!

1. Fork the repository.
2. Create a new branch (`git checkout -b feature-branch`).
3. Commit your changes (`git commit -am 'Add new feature'`).
4. Push to the branch (`git push origin feature-branch`).
5. Open a pull request.

## License

This project is licensed and worked under O365 and Exchnage environment.

## Contact

For any queries or support, feel free to contact:

- **Your Name**:-Kirti Ranjan Behera
- **Email**: Kirtiranjan1988@gmail.com

---

You can modify this template further based on specific details of your script and usage scenarios!
