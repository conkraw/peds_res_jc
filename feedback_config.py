"""Fixed feedback links for journal club exports."""

THANK_YOU_TITLE = "Thank you for participating"
THANK_YOU_MESSAGE = (
    "Please complete the brief REDCap feedback form so we can keep journal club "
    "clinically relevant, relaxed, and useful."
)
FEEDBACK_INSTRUCTION = "Scan the QR code or use the link to provide feedback."

# The QR code always points to the full REDCap survey URL.
REDCAP_QR_URL = "https://redcap.ctsi.psu.edu/surveys/?s=T9P4FPRYMJ3XL478"

# The displayed website link is the short REDCap link for laptop users.
REDCAP_DISPLAY_URL = "https://redcap.link/peds_res_jc_feedback"
