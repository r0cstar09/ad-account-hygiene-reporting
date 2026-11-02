# From a review queue to an approved account change

The workbook supports triage. Actual disablement, credential changes and deletion are intentionally outside this project.

1. **Confirm the evidence.** Check run time, scope, DC coverage and the audit sidecar. Investigate recent-activity or unknown-value rows. Requery if the directory changed or an object was missing on a replica.
2. **Establish identity and ownership.** Compare workforce accounts with authoritative HR/identity records. Identify the accountable owner for application and service identities. A name or OU is not sufficient evidence.
3. **Verify actual privilege.** Review group memberships, delegated permissions and other access paths. The `Privileged Accounts` sheet is an OU/naming classification, not an entitlement calculation.
4. **Check dependencies and exceptions.** Review services, scheduled tasks, integrations and relevant application logs with the owner. Explicitly check emergency-access, leave, seasonal and newly provisioned accounts. Do not apply a workforce inactivity policy mechanically to a non-human identity.
5. **Record a decision.** Track reviewer, evidence, owner response, approved action, ticket/change ID, due date and exception expiry in your authorized system of record. The generated workbook does not implement that approval workflow.
6. **Use a reversible change where policy permits.** Follow approved disable/quarantine procedures, monitor impacts and preserve a tested recovery path. Do not turn a report into a bulk deletion list.
7. **Verify and close.** Check the directory and dependent services after the approved action, rerun the report, reconcile identities rather than merely comparing totals, and retain evidence according to policy.

The useful outcome is an explainable decision per account—not the largest possible deletion count. Do not claim risk reduction, hours saved or a completed remediation campaign from a spreadsheet demonstration alone.
