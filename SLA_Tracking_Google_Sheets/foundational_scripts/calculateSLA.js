// ===================================================================
// SLA NA CONFIGURATION
// ===================================================================
// "Comment" dropdown options that mark a delay as Molina-caused.
// If a row's Comment matches one of these, any NA result becomes
// "NA - Molina related issue"; otherwise it becomes "NA - Auracom related issue".
// Keep these in sync with the dropdown range on the "ℹ️ HOW TO USE" tab.

const SLA_NA_CONFIG = {
  MOLINA_COMMENT_OPTIONS: [
    "Complex data processing - Molina related issue",
    "Molina delay to notify language needed - Molina related issue"
  ],
  NA_MOLINA: "NA - Molina related issue",
  NA_AURACOM: "NA - Auracom related issue"
};

/**
 * Determine which NA label to return based on the row's Comment value.
 * Matching ignores leading/trailing whitespace and letter case.
 *
 * @param {string} comment_column The value in the row's "Comment" column.
 * @return {string} "NA - Molina related issue" or "NA - Auracom related issue".
 */
function calculateCorrectNA_return(comment_column) {
  const comment = String(comment_column || "").trim().toLowerCase();

  const isMolinaIssue = SLA_NA_CONFIG.MOLINA_COMMENT_OPTIONS.some(
    option => option.trim().toLowerCase() === comment
  );

  return isMolinaIssue ? SLA_NA_CONFIG.NA_MOLINA : SLA_NA_CONFIG.NA_AURACOM;
}

/**
 * Calculate SLA Compliance based on the output of calculateStatus.gs.
 * * Determines if a file delivery meets SLA requirements by leveraging the
 * pre-calculated hours from the 'status_column'.
 *
 * @param {string} template_type The template type: "Word" or "PDF".
 * @param {string} isUrgent_column "YES" or "NO" (only used for Word template).
 * @param {string} status_column The output string from calculateStatus (e.g., "10.0 Business Hours").
 * @param {string} comment_column The row's "Comment" value (optional; used only to label NA results).
 * @return {string} "Y" if SLA met, "N" if SLA not met, or
 *                  "NA - Molina related issue" / "NA - Auracom related issue" if data is missing/invalid.
 * @customfunction
 */
function calculateSLA(template_type, isUrgent_column, status_column, comment_column) {

    // ===================================================================
    // STEP 1: Validate Input and Standardize
    // ===================================================================
    if (!template_type || !status_column) {
      return calculateCorrectNA_return(comment_column);
    }

    const template = String(template_type).trim().toUpperCase();
    const status = String(status_column).trim();
    const isUrgent = String(isUrgent_column || "").trim().toUpperCase();

    // ===================================================================
    // STEP 2: Parse Hours from the Status String
    // ===================================================================
    // Regex to find the first sequence of digits, optionally followed by a decimal point and more digits.
    const hourMatch = status.match(/(\d+\.?\d*)/);

    if (!hourMatch || hourMatch.length < 2) {
      // THIS IS THE CRITICAL LOGIC: If no number is found, it returns an NA label.
      return calculateCorrectNA_return(comment_column);
    }

    const hour_difference = parseFloat(hourMatch[1]);

    // Also catches strings like "Inf" or other non-standard numeric returns from the regex.
    if (isNaN(hour_difference)) {
      return calculateCorrectNA_return(comment_column);
    }

    // ===================================================================
    // STEP 3: Apply SLA Compliance Rules
    // ===================================================================

    let sla_threshold; // In hours

    switch (template) {
      case "WORD":
        // Rule 1: Word Urgent (YES) - 6 hours
        if (isUrgent === "YES") {
          sla_threshold = 6;
        }
        // Rule 2: Word Regular (NO) - 24 hours
        else if (isUrgent === "NO") {
          sla_threshold = 24;
        }
        // Rule 3: Word with Invalid/Missing Urgency - Return NA label
        else {
          return calculateCorrectNA_return(comment_column);
        }
        break;

      case "PDF":
        // Rule 4: PDF - 60 hours
        sla_threshold = 60;
        break;

      default:
        // Invalid Template Type
        return calculateCorrectNA_return(comment_column);
    }

    // ===================================================================
    // STEP 4: Final Compliance Check
    // ===================================================================

    // Return "Y" if the total hours are less than or equal to the required threshold.
    return hour_difference <= sla_threshold ? "Y" : "N";
  }