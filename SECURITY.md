# Security and macro trust

TerminPlanner is an offline Excel workbook. Its VBA source is available in
`src/VBA` and embedded into the released `.xlsm` by the public GitHub Actions
workflow.

The release is not digitally signed. Excel may block macros in files downloaded
from the internet. Inspect the source and release checksums first. On Windows,
right-click the downloaded file, open **Properties**, select **Unblock**, apply,
then open it in desktop Excel. Enable macros only for a copy you trust.

TerminPlanner does not upload data. Appointments remain inside the workbook;
the optional export writes a CSV beside the saved workbook. Neither file is
encrypted, so protect them with appropriate operating-system permissions and
backups.
