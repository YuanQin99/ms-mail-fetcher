<script setup>
import { computed, onMounted, ref } from 'vue'
import { useRoute, useRouter } from 'vue-router'
import { ElMessage, ElMessageBox } from 'element-plus'
import { CopyDocument } from '@element-plus/icons-vue'
import {
    archiveAccount,
    archiveAllAccounts,
    createAccountType,
    deleteAccountType,
    exportAccountsUrl,
    exportSelectedAccountsUrl,
    getAccounts,
    getAccountTypes,
    importAccounts,
    refreshAllAccountTokens,
    updateAccount,
    updateAccountType,
} from '@/api/accounts'

const router = useRouter()
const route = useRoute()

const loading = ref(false)
const importLoading = ref(false)
const refreshAllLoading = ref(false)
const rows = ref([])
const total = ref(0)
const page = ref(1)
const pageSize = ref(10)
const search = ref('')
const accountType = ref('')
const accountTypes = ref([])

const pasteDialogVisible = ref(false)
const pasteContent = ref('')
const selectedRows = ref([])
const batchType = ref('')
const batchTypeDialogVisible = ref(false)
const batchTypeSaving = ref(false)
const batchRemark = ref('')
const batchRemarkDialogVisible = ref(false)
const batchRemarkSaving = ref(false)
const typeDialogVisible = ref(false)
const typeCreateForm = ref({ code: '', label: '', color: '#409EFF' })
const savingType = ref(false)
const typeEditCache = ref({})
const detailDialogVisible = ref(false)
const detailEditMode = ref(false)
const editSaving = ref(false)
const editForm = ref({
    id: null,
    email: '',
    account_type: '',
    remark: '',
})

const fileInputRef = ref(null)

const thresholdDays = 30

function emitAccountsChanged() {
    if (typeof window !== 'undefined') {
        window.dispatchEvent(new Event('accounts-changed'))
    }
}

const handleCopy = (text) => {
    if (!text) return
    navigator.clipboard.writeText(text).then(() => {
        ElMessage.success('已复制')
    }).catch(() => {
        ElMessage.error('复制失败')
    })
}

const pageCount = computed(() => Math.max(Math.ceil(total.value / pageSize.value), 1))
const typeMap = computed(() => {
    const map = {}
    for (const item of accountTypes.value) {
        map[item.code] = item
    }
    return map
})

async function fetchAccountTypes() {
    try {
        const data = await getAccountTypes()
        accountTypes.value = Array.isArray(data) ? data : []
        typeEditCache.value = {}
        for (const item of accountTypes.value) {
            typeEditCache.value[item.id] = {
                label: item.label,
                color: item.color,
            }
        }
    } catch (error) {
        ElMessage.error(error.message || '加载账号类型失败')
    }
}

async function fetchData() {
    loading.value = true
    try {
        const data = await getAccounts({
            is_active: true,
            search: search.value,
            type: accountType.value,
            page: page.value,
            page_size: pageSize.value,
        })
        rows.value = data.items || []
        total.value = data.total || 0
        emitAccountsChanged()
    } catch (error) {
        ElMessage.error(error.message || '加载失败')
    } finally {
        loading.value = false
    }
}

function toPositiveInt(value, fallback) {
    const n = Number.parseInt(String(value ?? ''), 10)
    return Number.isFinite(n) && n > 0 ? n : fallback
}

function applyStateFromRouteQuery() {
    page.value = toPositiveInt(route.query.page, 1)
    pageSize.value = toPositiveInt(route.query.page_size, 10)
    search.value = String(route.query.search || '')
    accountType.value = String(route.query.type || '')
}

function syncRouteQuery() {
    const query = {
        page: String(page.value),
        page_size: String(pageSize.value),
    }
    if (search.value) query.search = search.value
    if (accountType.value) query.type = accountType.value

    router.replace({
        path: '/active',
        query,
    })
}

function onSearch() {
    page.value = 1
    syncRouteQuery()
    fetchData()
}

function resetFilters() {
    search.value = ''
    accountType.value = ''
    page.value = 1
    syncRouteQuery()
    fetchData()
}

function goMail(row, folder) {
    router.push({
        path: `/mail/${row.id}/${folder}`,
        query: {
            email: row.email,
            is_active: 'true',
            _from: '/active',
            _page: String(page.value),
            _page_size: String(pageSize.value),
            _search: search.value,
            _type: accountType.value,
        },
    })
}

function onPageChange() {
    syncRouteQuery()
    fetchData()
}

async function onArchive(row) {
    try {
        await ElMessageBox.confirm(`确认归档账号 ${row.email} 吗？`, '提示', {
            type: 'warning',
            confirmButtonText: '确认',
            cancelButtonText: '取消',
        })
        await archiveAccount(row.id)
        ElMessage.success('已归档')
        if (rows.value.length === 1 && page.value > 1) page.value -= 1
        fetchData()
    } catch {
        // canceled
    }
}

async function onRefreshAllTokens() {
    try {
        await ElMessageBox.confirm('确认一键刷新当前筛选范围内所有账号 Token 吗？', '提示', {
            type: 'warning',
            confirmButtonText: '确认刷新',
            cancelButtonText: '取消',
        })

        refreshAllLoading.value = true
        const result = await refreshAllAccountTokens({
            is_active: true,
            search: search.value,
            type: accountType.value,
        })
        ElMessage.success(`刷新完成：总数 ${result.total || 0}，成功 ${result.success || 0}，失败 ${result.failed || 0}`)
        await fetchData()
    } catch (error) {
        if (error?.message) {
            ElMessage.error(error.message)
        }
    } finally {
        refreshAllLoading.value = false
    }
}

async function onArchiveAll() {
    try {
        await ElMessageBox.confirm('确认将当前所有活跃账号归档吗？', '危险操作', {
            type: 'warning',
            confirmButtonText: '确认归档',
            cancelButtonText: '取消',
        })
        await archiveAllAccounts()
        ElMessage.success('全部归档完成')
        page.value = 1
        fetchData()
    } catch {
        // canceled
    }
}

function onSelectionChange(selection) {
    selectedRows.value = selection
}

async function onBatchArchiveSelected() {
    if (!selectedRows.value.length) {
        ElMessage.warning('请先勾选账号')
        return
    }

    try {
        await ElMessageBox.confirm(
            `确认归档已选中的 ${selectedRows.value.length} 个账号吗？`,
            '批量操作确认',
            {
                type: 'warning',
                confirmButtonText: '确认归档',
                cancelButtonText: '取消',
            },
        )

        const results = await Promise.allSettled(
            selectedRows.value.map((row) => archiveAccount(row.id)),
        )
        const successCount = results.filter((r) => r.status === 'fulfilled').length
        const failedCount = results.length - successCount

        if (failedCount > 0) {
            ElMessage.warning(`批量归档完成：成功 ${successCount}，失败 ${failedCount}`)
        } else {
            ElMessage.success(`批量归档完成：共 ${successCount} 条`)
        }

        if (rows.value.length === successCount && page.value > 1) page.value -= 1
        selectedRows.value = []
        fetchData()
    } catch {
        // canceled
    }
}

async function onBatchUpdateType() {
    if (!selectedRows.value.length) {
        ElMessage.warning('请先勾选账号')
        return
    }
    if (!batchType.value) {
        ElMessage.warning('请先选择要批量设置的账号类型')
        return
    }

    batchTypeSaving.value = true
    try {
        await ElMessageBox.confirm(
            `确认将已选 ${selectedRows.value.length} 个账号更新为类型 ${getTypeLabel(batchType.value)} 吗？`,
            '批量修改确认',
            {
                type: 'warning',
                confirmButtonText: '确认修改',
                cancelButtonText: '取消',
            },
        )

        const results = await Promise.allSettled(
            selectedRows.value.map((row) => updateAccount(row.id, { account_type: batchType.value }, { is_active: true })),
        )
        const successCount = results.filter((r) => r.status === 'fulfilled').length
        const failedCount = results.length - successCount

        if (failedCount > 0) {
            ElMessage.warning(`批量修改完成：成功 ${successCount}，失败 ${failedCount}`)
        } else {
            ElMessage.success(`批量修改完成：共 ${successCount} 条`)
        }

        selectedRows.value = []
        batchType.value = ''
        batchTypeDialogVisible.value = false
        await fetchData()
    } catch {
        // canceled
    } finally {
        batchTypeSaving.value = false
    }
}

function onOpenBatchTypeDialog() {
    if (!selectedRows.value.length) {
        ElMessage.warning('请先勾选账号')
        return
    }
    batchType.value = ''
    batchTypeDialogVisible.value = true
}

function onOpenBatchRemarkDialog() {
    if (!selectedRows.value.length) {
        ElMessage.warning('请先勾选账号')
        return
    }
    batchRemark.value = ''
    batchRemarkDialogVisible.value = true
}

async function onBatchUpdateRemark() {
    if (!selectedRows.value.length) {
        ElMessage.warning('请先勾选账号')
        return
    }

    batchRemarkSaving.value = true
    try {
        await ElMessageBox.confirm(
            `确认将已选 ${selectedRows.value.length} 个账号的备注批量修改为当前内容吗？`,
            '批量备注确认',
            {
                type: 'warning',
                confirmButtonText: '确认修改',
                cancelButtonText: '取消',
            },
        )

        const remarkValue = batchRemark.value.trim()
        const results = await Promise.allSettled(
            selectedRows.value.map((row) => updateAccount(row.id, { remark: remarkValue || null }, { is_active: true })),
        )
        const successCount = results.filter((r) => r.status === 'fulfilled').length
        const failedCount = results.length - successCount

        if (failedCount > 0) {
            ElMessage.warning(`批量备注完成：成功 ${successCount}，失败 ${failedCount}`)
        } else {
            ElMessage.success(`批量备注完成：共 ${successCount} 条`)
        }

        selectedRows.value = []
        batchRemark.value = ''
        batchRemarkDialogVisible.value = false
        await fetchData()
    } catch {
        // canceled
    } finally {
        batchRemarkSaving.value = false
    }
}

function onOpenDetail(row) {
    detailEditMode.value = false
    editForm.value = {
        id: row.id,
        email: row.email,
        account_type: row.account_type || '',
        remark: row.remark || '',
    }
    detailDialogVisible.value = true
}

function onStartEdit() {
    detailEditMode.value = true
}

function onCancelDetailEdit() {
    detailEditMode.value = false
}

async function onSaveEdit() {
    if (!editForm.value.id) return

    editSaving.value = true
    try {
        await updateAccount(editForm.value.id, {
            account_type: editForm.value.account_type || null,
            remark: editForm.value.remark || null,
        }, { is_active: true })
        ElMessage.success('账号信息已更新')
        detailEditMode.value = false
        await fetchData()
    } catch (error) {
        ElMessage.error(error.message || '保存失败')
    } finally {
        editSaving.value = false
    }
}

function getTypeTagType(code) {
    if (!code) return 'info'
    const lower = String(code).toLowerCase()
    if (lower === 'team') return 'primary'
    if (lower === 'member') return 'success'
    if (lower === 'plus') return 'warning'
    if (lower === 'idle') return 'info'
    return 'info'
}

function getTypeLabel(code) {
    return typeMap.value[code]?.label || code || '-'
}

function onOpenTypeDialog() {
    typeDialogVisible.value = true
    fetchAccountTypes()
}

async function onCreateType() {
    if (!typeCreateForm.value.code.trim()) {
        ElMessage.warning('请输入类型编码')
        return
    }
    if (!typeCreateForm.value.label.trim()) {
        ElMessage.warning('请输入类型名称')
        return
    }

    savingType.value = true
    try {
        await createAccountType({
            code: typeCreateForm.value.code.trim().toLowerCase(),
            label: typeCreateForm.value.label.trim(),
            color: typeCreateForm.value.color,
        })
        ElMessage.success('类型已添加')
        typeCreateForm.value = { code: '', label: '', color: '#409EFF' }
        await fetchAccountTypes()
    } catch (error) {
        ElMessage.error(error.message || '添加类型失败')
    } finally {
        savingType.value = false
    }
}

async function onSaveType(item) {
    const edit = typeEditCache.value[item.id]
    if (!edit?.label?.trim()) {
        ElMessage.warning('类型名称不能为空')
        return
    }

    try {
        await updateAccountType(item.id, {
            label: edit.label.trim(),
            color: edit.color,
        })
        ElMessage.success('类型已更新')
        await fetchAccountTypes()
        fetchData()
    } catch (error) {
        ElMessage.error(error.message || '更新类型失败')
    }
}

async function onDeleteType(item) {
    try {
        await ElMessageBox.confirm(
            `确认删除类型 ${item.label}(${item.code}) 吗？已绑定账号将清空类型。`,
            '提示',
            {
                type: 'warning',
                confirmButtonText: '确认删除',
                cancelButtonText: '取消',
            },
        )
        await deleteAccountType(item.id)
        ElMessage.success('类型已删除')
        if (accountType.value === item.code) {
            accountType.value = ''
            page.value = 1
            fetchData()
        }
        await fetchAccountTypes()
    } catch {
        // canceled
    }
}

function onExportSelected() {
    if (!selectedRows.value.length) {
        ElMessage.warning('请先勾选账号')
        return
    }

    const url = exportSelectedAccountsUrl(
        {
            is_active: true,
            search: search.value,
            type: accountType.value,
        },
        selectedRows.value.map((row) => row.id),
        'accounts_selected_active',
    )
    window.open(url, '_blank')
}

function onExport() {
    const url = exportAccountsUrl({
        is_active: true,
        search: search.value,
        type: accountType.value,
    }, 'accounts_export_active')
    window.open(url, '_blank')
}

function onClickFileImport() {
    fileInputRef.value?.click()
}

async function onFileSelected(event) {
    const file = event.target.files?.[0]
    if (!file) return

    importLoading.value = true
    try {
        const formData = new FormData()
        formData.append('file', file)
        formData.append('is_active', 'true')
        if (accountType.value) formData.append('account_type', accountType.value)

        const result = await importAccounts(formData)
        ElMessage.success(`导入完成：新增 ${result.inserted}，跳过 ${result.skipped}`)
        if (result.errors?.length) {
            ElMessage.warning(`有 ${result.errors.length} 行格式错误，请检查输入`)
        }
        page.value = 1
        fetchData()
    } catch (error) {
        ElMessage.error(error.message || '文件导入失败')
    } finally {
        importLoading.value = false
        event.target.value = ''
    }
}

async function onPasteImport() {
    if (!pasteContent.value.trim()) {
        ElMessage.warning('请输入导入文本')
        return
    }

    importLoading.value = true
    try {
        const formData = new FormData()
        formData.append('text', pasteContent.value)
        formData.append('is_active', 'true')
        if (accountType.value) formData.append('account_type', accountType.value)

        const result = await importAccounts(formData)
        ElMessage.success(`导入完成：新增 ${result.inserted}，跳过 ${result.skipped}`)
        if (result.errors?.length) {
            ElMessage.warning(`有 ${result.errors.length} 行格式错误，请检查输入`)
        }
        pasteDialogVisible.value = false
        pasteContent.value = ''
        page.value = 1
        fetchData()
    } catch (error) {
        ElMessage.error(error.message || '粘贴导入失败')
    } finally {
        importLoading.value = false
    }
}

onMounted(async () => {
    applyStateFromRouteQuery()
    await fetchAccountTypes()
    fetchData()
})
</script>

<template>
    <div class="page-container">
        <el-card shadow="never" class="page-card premium-card" >
            <template #header>
                <div class="toolbar">
                    <div class="left">
                        <el-input v-model="search" placeholder="搜索 email / remark" clearable style="width: 240px"
                            @keyup.enter="onSearch" />
                        <el-select v-model="accountType" placeholder="账号类型" clearable style="width: 150px">
                            <el-option v-for="item in accountTypes" :key="item.id" :label="item.label"
                                :value="item.code" />
                        </el-select>
                        <el-button type="primary" @click="onSearch">查询</el-button>
                        <el-button @click="resetFilters">重置</el-button>
                        <el-button @click="onOpenTypeDialog">类型管理</el-button>
                    </div>

                    <div class="right">
                        <input ref="fileInputRef" type="file" accept=".txt" class="hidden" @change="onFileSelected" />
                        <el-button :loading="importLoading" @click="onClickFileImport">文件导入</el-button>
                        <el-button :loading="importLoading" @click="pasteDialogVisible = true">粘贴导入</el-button>
                        <el-button @click="onExport">一键导出</el-button>
                        <el-button type="primary" plain :disabled="!selectedRows.length" @click="onExportSelected">
                            批量导出已选（{{ selectedRows.length }}）
                        </el-button>
                        <el-button :loading="refreshAllLoading" @click="onRefreshAllTokens">一键刷新Token</el-button>
                        <el-button type="primary" plain :disabled="!selectedRows.length" @click="onOpenBatchTypeDialog">
                            批量改类型（{{ selectedRows.length }}）
                        </el-button>
                        <el-button type="primary" plain :disabled="!selectedRows.length" @click="onOpenBatchRemarkDialog">
                            批量改备注（{{ selectedRows.length }}）
                        </el-button>
                        <el-button type="warning" plain :disabled="!selectedRows.length"
                            @click="onBatchArchiveSelected">
                            批量归档已选（{{ selectedRows.length }}）
                        </el-button>
                        <el-button type="danger" plain @click="onArchiveAll">清空活跃池</el-button>
                    </div>
                </div>
            </template>

            <el-table v-loading="loading" :data="rows" border stripe @selection-change="onSelectionChange"
                style="flex: 1; height: 100%; min-height: 0;">
                <el-table-column type="selection" width="48" />
                <el-table-column label="序号" width="80" align="center">
                    <template #default="{ $index }">
                        {{ (page - 1) * pageSize + $index + 1 }}
                    </template>
                </el-table-column>
                <el-table-column label="邮箱" min-width="200">
                    <template #default="{ row }">
                        <div class="copy-cell">
                            <span class="truncate-text">{{ row.email }}</span>
                            <el-button link type="primary" :icon="CopyDocument" @click="handleCopy(row.email)" />
                        </div>
                    </template>
                </el-table-column>
                <el-table-column label="密码" min-width="160">
                    <template #default="{ row }">
                        <div class="copy-cell">
                            <span class="truncate-text">{{ row.password }}</span>
                            <el-button link type="primary" :icon="CopyDocument" @click="handleCopy(row.password)" />
                        </div>
                    </template>
                </el-table-column>
                <el-table-column label="距上次刷新天数" width="140" align="center">
                    <template #default="{ row }">
                        <el-tag :type="row.days_since_refresh > thresholdDays ? 'danger' : 'success'" effect="light"
                            round>
                            {{ row.days_since_refresh }} 天
                        </el-tag>
                    </template>
                </el-table-column>
                <el-table-column label="账号类型" width="120" align="center" header-align="center">
                    <template #default="{ row }">
                        <el-tag :type="getTypeTagType(row.account_type)" :color="typeMap[row.account_type]?.color"
                            effect="dark">
                            {{ getTypeLabel(row.account_type) }}
                        </el-tag>
                    </template>
                </el-table-column>
                <el-table-column label="备注" min-width="140">
                    <template #default="{ row }">
                        <span class="remark-ellipsis">{{ row.remark || '' }}</span>
                    </template>
                </el-table-column>
                <el-table-column label="操作" width="240" fixed="right">
                    <template #default="{ row }">
                        <el-space>
                            <el-button type="primary" link @click="goMail(row, 'inbox')">收信</el-button>
                            <el-button type="warning" link @click="goMail(row, 'spam')">垃圾</el-button>
                            <el-button type="success" link @click="onOpenDetail(row)">详情</el-button>
                            <el-button type="danger" link @click="onArchive(row)">归档</el-button>
                        </el-space>
                    </template>
                </el-table-column>
            </el-table>

            <div class="pager">
                <el-pagination v-model:current-page="page" v-model:page-size="pageSize"
                    layout="total, sizes, prev, pager, next, jumper" :total="total" :page-count="pageCount"
                    :page-sizes="[10, 20, 50, 100]" @change="onPageChange" />
            </div>
        </el-card>

        <el-dialog v-model="pasteDialogVisible" title="粘贴导入" width="680px">
            <el-alert type="info" show-icon :closable="false" title="每行格式：邮箱----密码----client_id----refresh_token"
                class="mb-12" />
            <el-input v-model="pasteContent" type="textarea" :rows="10" placeholder="按行粘贴账号数据" />
            <template #footer>
                <el-button @click="pasteDialogVisible = false">取消</el-button>
                <el-button type="primary" :loading="importLoading" @click="onPasteImport">开始导入</el-button>
            </template>
        </el-dialog>

        <el-dialog v-model="typeDialogVisible" title="账号类型管理" width="780px">
            <el-alert type="info" show-icon :closable="false" title="类型编码用于账号绑定与筛选，颜色会持久化保存到数据库。" class="mb-12" />

            <el-card shadow="never" class="mb-12">
                <div class="type-create-row">
                    <el-input v-model="typeCreateForm.code" placeholder="编码，例如 team / member / plus / idle"
                        style="width: 200px" />
                    <el-input v-model="typeCreateForm.label" placeholder="显示名称，例如 Team / Member" style="width: 200px" />
                    <el-color-picker v-model="typeCreateForm.color" />
                    <el-button type="primary" :loading="savingType" @click="onCreateType">添加类型</el-button>
                </div>
            </el-card>

            <el-table :data="accountTypes" border stripe>
                <el-table-column prop="code" label="编码" width="140" />
                <el-table-column label="名称" min-width="180">
                    <template #default="{ row }">
                        <el-input v-model="typeEditCache[row.id].label" />
                    </template>
                </el-table-column>
                <el-table-column label="颜色" width="120" align="center">
                    <template #default="{ row }">
                        <el-color-picker v-model="typeEditCache[row.id].color" />
                    </template>
                </el-table-column>
                <el-table-column label="预览" width="120" align="center">
                    <template #default="{ row }">
                        <el-tag :color="typeEditCache[row.id].color" effect="dark">{{ typeEditCache[row.id].label }}</el-tag>
                    </template>
                </el-table-column>
                <el-table-column label="操作" width="180" align="center">
                    <template #default="{ row }">
                        <el-space>
                            <el-button type="primary" link @click="onSaveType(row)">保存</el-button>
                            <el-button type="danger" link @click="onDeleteType(row)">删除</el-button>
                        </el-space>
                    </template>
                </el-table-column>
            </el-table>
        </el-dialog>

        <el-dialog v-model="detailDialogVisible" :title="detailEditMode ? '编辑账号' : '账号详情'" width="520px">
            <el-form label-width="90px">
                <el-form-item label="邮箱">
                    <el-input :model-value="editForm.email" disabled />
                </el-form-item>
                <el-form-item label="账号类型">
                    <el-select v-model="editForm.account_type" clearable placeholder="请选择账号类型" style="width: 100%" :disabled="!detailEditMode">
                        <el-option v-for="item in accountTypes" :key="`edit-${item.id}`" :label="item.label"
                            :value="item.code" />
                    </el-select>
                </el-form-item>
                <el-form-item label="备注">
                    <el-input v-model="editForm.remark" type="textarea" :rows="3" placeholder="请输入备注" :disabled="!detailEditMode" />
                </el-form-item>
            </el-form>
            <template #footer>
                <el-button @click="detailDialogVisible = false">关闭</el-button>
                <el-button v-if="!detailEditMode" type="primary" @click="onStartEdit">编辑</el-button>
                <el-button v-else @click="onCancelDetailEdit">取消编辑</el-button>
                <el-button v-if="detailEditMode" type="primary" :loading="editSaving" @click="onSaveEdit">保存</el-button>
            </template>
        </el-dialog>

        <el-dialog v-model="batchTypeDialogVisible" title="批量修改账号类型" width="460px">
            <el-alert type="info" :closable="false" show-icon :title="`将修改已选中的 ${selectedRows.length} 个账号`"
                class="mb-12" />
            <el-form label-width="90px">
                <el-form-item label="目标类型">
                    <el-select v-model="batchType" clearable placeholder="请选择账号类型" style="width: 100%">
                        <el-option v-for="item in accountTypes" :key="`batch-${item.id}`" :label="item.label"
                            :value="item.code" />
                    </el-select>
                </el-form-item>
            </el-form>
            <template #footer>
                <el-button @click="batchTypeDialogVisible = false">取消</el-button>
                <el-button type="primary" :loading="batchTypeSaving" @click="onBatchUpdateType">确认修改</el-button>
            </template>
        </el-dialog>

        <el-dialog v-model="batchRemarkDialogVisible" title="批量修改备注" width="520px">
            <el-alert type="info" :closable="false" show-icon :title="`将修改已选中的 ${selectedRows.length} 个账号备注`"
                class="mb-12" />
            <el-form label-width="90px">
                <el-form-item label="备注">
                    <el-input v-model="batchRemark" type="textarea" :rows="4" placeholder="留空则清空备注" />
                </el-form-item>
            </el-form>
            <template #footer>
                <el-button @click="batchRemarkDialogVisible = false">取消</el-button>
                <el-button type="primary" :loading="batchRemarkSaving" @click="onBatchUpdateRemark">确认修改</el-button>
            </template>
        </el-dialog>
    </div>
</template>

<style scoped>
.page-card {
    flex: 1;
    display: flex;
    flex-direction: column;
    min-height: 600px;
    border: none;
}

:deep(.el-card__body) {
    flex: 1;
    display: flex;
    flex-direction: column;
    min-height: 0;
    overflow: hidden;
    padding-bottom: 0px;
}

.toolbar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    flex-wrap: wrap;
}

.left,
.right {
    display: flex;
    align-items: center;
    gap: 8px;
    flex-wrap: wrap;
}

.hidden {
    display: none;
}

.warn {
    color: #f56c6c;
    font-weight: 700;
}

.pager {
    margin-top: 16px;
    display: flex;
    justify-content: flex-end;
    margin-bottom: 20px;
}

.mb-12 {
    margin-bottom: 12px;
}

.type-create-row {
    display: flex;
    align-items: center;
    gap: 12px;
    flex-wrap: wrap;
}

.copy-cell {
    display: flex;
    align-items: center;
    gap: 8px;
    width: 100%;
}

.truncate-text {
    flex: 1;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}

.remark-ellipsis {
    display: inline-block;
    max-width: 100%;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    vertical-align: middle;
}

.refresh-btn {
    background: linear-gradient(135deg, #67c23a, #85ce61);
    border: none;
    transition: transform 0.2s, box-shadow 0.2s;
}

.refresh-btn:hover {
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(103, 194, 58, 0.4);
}
</style>
