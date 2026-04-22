import { buildUrl, request } from './http'

export function getAccounts(params) {
  return request('/api/accounts', {}, params)
}

export function createAccount(payload) {
  return request('/api/accounts', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
}

export function updateAccount(id, payload, params = {}) {
  return request(`/api/accounts/${id}`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  }, params)
}

export function archiveAccount(id) {
  return request(`/api/accounts/${id}/archive`, { method: 'PUT' })
}

export function refreshAllAccountTokens(params) {
  return request('/api/accounts/refresh-all-tokens', { method: 'PUT' }, params)
}

export function archiveAllAccounts() {
  return request('/api/accounts/archive-all', { method: 'PUT' })
}

export function deleteAccount(id, params = {}) {
  return request(`/api/accounts/${id}`, { method: 'DELETE' }, params)
}

export function importAccounts(formData) {
  return request('/api/accounts/import', {
    method: 'POST',
    body: formData,
  })
}

export function exportAccountsUrl(params, filenamePrefix) {
  return buildUrl('/api/accounts/export', { ...params, filename_prefix: filenamePrefix })
}

export function exportSelectedAccountsUrl(params, ids, filenamePrefix) {
  const accountIds = Array.isArray(ids) ? ids.filter(Boolean).join(',') : ''
  return buildUrl('/api/accounts/export', { ...params, ids: accountIds, filename_prefix: filenamePrefix })
}

export function getMailList(accountId, folder, params) {
  return request(`/api/accounts/${accountId}/mail/${folder}`, {}, params)
}

export function getMailDetail(accountId, folder, messageId, params) {
  return request(`/api/accounts/${accountId}/mail/${folder}/${messageId}`, {}, params)
}

export function getAccountTypes() {
  return request('/api/account-types')
}

export function createAccountType(payload) {
  return request('/api/account-types', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
}

export function updateAccountType(id, payload) {
  return request(`/api/account-types/${id}`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
}

export function deleteAccountType(id) {
  return request(`/api/account-types/${id}`, { method: 'DELETE' })
}

export function getUiPreferences() {
  return request('/api/ui/preferences')
}

export function updateUiPreferences(payload) {
  return request('/api/ui/preferences', {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
}
