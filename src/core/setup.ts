/**
 * Guided setup logic (pure parts). Graph refuses consumer-identity app
 * creation via API (ADR-0002), so setup walks the user through the Entra
 * portal, then writes env values and verifies the registration via the normal
 * device-code auth. Step numbers here must match the "Manual setup" section
 * of the README — verification errors refer back to them.
 */

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export function isValidClientId(value: string): boolean {
  return GUID_RE.test(value);
}

export function buildEnvUpdates(
  clientId: string,
  personalOnly: boolean,
): Record<string, string | null> {
  return {
    OUTLOOK_MAIL_CLIENT_ID: clientId,
    // Tenant is derived from the audience choice, never set independently —
    // this makes the AADSTS9002331 audience/tenant mismatch unrepresentable.
    // Default mode must also *remove* a stale consumers value left by an
    // earlier --personal-only run (null deletes the key).
    OUTLOOK_MAIL_TENANT_ID: personalOnly ? "consumers" : null,
  };
}

export function buildSetupSteps(personalOnly: boolean): string {
  const audience = personalOnly
    ? "Personal Microsoft accounts only"
    : "Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant) and personal Microsoft accounts";
  return [
    "Register the app in the Entra portal (Microsoft blocks doing this automatically for personal accounts):",
    "",
    "  1. Open https://entra.microsoft.com -> Identity -> Applications -> App registrations -> New registration.",
    "  2. Name: anything (e.g. outlook-mail).",
    `  3. Supported account types: choose "${audience}".`,
    "  4. Redirect URI: leave blank. Click Register.",
    "  5. On the Overview page, copy the Application (client) ID.",
    "  6. Authentication -> Advanced settings -> Allow public client flows -> Yes -> Save.",
    "  7. API permissions -> Add a permission -> Microsoft Graph -> Delegated -> check Mail.Read and offline_access -> Add permissions.",
    "",
  ].join("\n");
}

const MANUAL_FALLBACK =
  'See the "Manual setup" section of the README, then re-run outlook-mail setup.';

/**
 * Maps a device-code verification failure to "go back and fix step N"
 * guidance. Input is the raw error message (AADSTS codes appear inline).
 */
export function translateVerifyError(message: string, personalOnly: boolean): string {
  if (message.includes("AADSTS9002331")) {
    return (
      "The app is registered for personal Microsoft accounts only, but setup ran in default (common) mode. " +
      "Either re-run setup with --personal-only, or change the account type chosen in step 3. " +
      MANUAL_FALLBACK
    );
  }
  if (message.includes("AADSTS50194")) {
    return (
      "The app is not registered for the consumers audience, but setup ran with --personal-only. " +
      "Either re-run setup without --personal-only, or change the account type chosen in step 3. " +
      MANUAL_FALLBACK
    );
  }
  if (message.includes("AADSTS7000218")) {
    return (
      'Public client flows are disabled on the app. Go back to step 6 and set "Allow public client flows" to Yes. ' +
      MANUAL_FALLBACK
    );
  }
  if (message.includes("AADSTS700016")) {
    return (
      `The client ID does not match any app registration${personalOnly ? " in the consumers tenant" : ""}. ` +
      "Check the value copied in step 5 (Application (client) ID on the Overview page, not the Object ID). " +
      MANUAL_FALLBACK
    );
  }
  if (message.includes("AADSTS65004")) {
    return (
      "You declined the consent prompt during sign-in. Re-run setup and accept the Mail.Read consent to finish. " +
      MANUAL_FALLBACK
    );
  }
  return `Verification failed: ${message}\n${MANUAL_FALLBACK}`;
}
