
function tc09TestSelfApproval() {
  const id = 'IAM-944c7bea-54fb-4132-a809-32c97a804e9e';
  const expectedRecord = 'AUD-02-2026';

  const access = rbacRequirePermission_('reviewResponses');

  const item = iamResponseRows_().find(function(row) {
    return row['Submission ID'] === id;
  });

  if (!item) throw new Error('STOP: Test submission not found.');

  if (item['Record ID'] !== expectedRecord ||
      item.Status !== 'Submitted' ||
      item['Submitted By'] !== access.email) {
    throw new Error(
      'STOP: Test record, status or reviewer identity mismatch.'
    );
  }

  let rejection = '';

  try {
    reviewIAMResponse_({
      id: id,
      decision: 'Accepted',
      comment: 'TC-09 backend self-approval security test'
    });
  } catch (error) {
    rejection = String(error.message || error);
  }

  const after = iamResponseRows_().find(function(row) {
    return row['Submission ID'] === id;
  });

  if (!after || after.Status !== 'Submitted') {
    throw new Error(
      'TC09 FAIL: Submission status changed unexpectedly.'
    );
  }

  if (rejection ===
      'Another Compliance reviewer must review your own submission.') {
    Logger.log('TC09_BACKEND_SELF_APPROVAL: PASS');
    Logger.log('Backend rejected self-approval.');
    Logger.log('Submission remains Submitted.');
    return;
  }

  throw new Error(
    'TC09 NOT PASSED: Expected rejection not received. ' +
    (rejection || 'No rejection')
  );
}
