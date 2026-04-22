<script setup>
import { computed, onMounted, ref, watch } from 'vue'
import { useRoute, useRouter } from 'vue-router'
import { ElMessage } from 'element-plus'
import { Refresh } from '@element-plus/icons-vue'
import { getMailDetail, getMailList } from '@/api/accounts'

const route = useRoute()
const router = useRouter()

const accountId = computed(() => Number(route.params.accountId))
const folder = computed(() => String(route.params.folder || 'inbox'))
const email = computed(() => String(route.query.email || `账号 #${accountId.value}`))
const isActive = computed(() => String(route.query.is_active || 'true') !== 'false')

const loading = ref(false)
const detailLoading = ref(false)

const rows = ref([])
const total = ref(0)
const page = ref(1)
const pageSize = ref(10)

const detailDialogVisible = ref(false)
const currentDetail = ref(null)

const pageCount = computed(() => Math.max(Math.ceil(total.value / pageSize.value), 1))
const folderText = computed(() => (folder.value === 'spam' ? '垃圾箱' : '收件箱'))

function looksLikeHtml(content) {
    if (!content) return false
    const text = String(content).trimStart().toLowerCase()
    return text.startsWith('<!doctype html') || text.startsWith('<html') || (text.includes('<body') && text.includes('</body>'))
}

const renderedBodyHtml = computed(() => {
    const detail = currentDetail.value
    if (!detail) return ''
    if (detail.body_html) return detail.body_html
    if (looksLikeHtml(detail.body_text)) return detail.body_text
    return ''
})

const renderedBodySrcdoc = computed(() => {
    const html = renderedBodyHtml.value
    if (!html) return ''

    const normalizedHtml = html.replace(/<a\b([^>]*?)>/gi, (match, attrs) => {
        let nextAttrs = attrs

        if (!/\btarget\s*=/i.test(nextAttrs)) {
            nextAttrs += ' target="_blank"'
        } else {
            nextAttrs = nextAttrs.replace(/\btarget\s*=\s*(['"]?)[^'"\s>]+\1?/i, 'target="_blank"')
        }

        if (!/\brel\s*=/i.test(nextAttrs)) {
            nextAttrs += ' rel="noopener noreferrer"'
        }

        return `<a${nextAttrs}>`
    })

    const lower = normalizedHtml.trimStart().toLowerCase()
    if (lower.startsWith('<!doctype html') || lower.startsWith('<html')) {
        return normalizedHtml
    }

    return `<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
      body {
        margin: 0;
        padding: 12px;
        color: #1f2937;
        font-family: Arial, sans-serif;
        word-break: break-word;
        line-height: 1.7;
        background: #fff;
      }
      img, table {
        max-width: 100%;
      }
      a {
        color: #2563eb;
      }
    </style>
  </head>
  <body>${normalizedHtml}</body>
</html>`
})

async function fetchMails() {
    if (!accountId.value) return

    loading.value = true
    try {
        const data = await getMailList(accountId.value, folder.value, {
            is_active: isActive.value,
            page: page.value,
            page_size: pageSize.value,
        })
        rows.value = data.items || []
        total.value = data.total || 0
    } catch (error) {
        ElMessage.error(error.message || '读取邮件列表失败')
    } finally {
        loading.value = false
    }
}

async function onViewDetail(row) {
    detailLoading.value = true
    detailDialogVisible.value = true
    try {
        const data = await getMailDetail(accountId.value, folder.value, row.uid, { is_active: isActive.value })
        currentDetail.value = data.detail
    } catch (error) {
        currentDetail.value = null
        ElMessage.error(error.message || '读取邮件详情失败')
    } finally {
        detailLoading.value = false
    }
}

function goBack() {
    const fromPath = String(route.query._from || '')

    if (fromPath === '/active' || fromPath === '/archived') {
        const query = {
            page: String(route.query._page || 1),
            page_size: String(route.query._page_size || 10),
        }

        const search = String(route.query._search || '')
        if (search) query.search = search

        const type = String(route.query._type || '')
        if (fromPath === '/active' && type) query.type = type

        router.push({
            path: fromPath,
            query,
        })
        return
    }

    if (window.history.length > 1) {
        router.back()
        return
    }

    router.push('/active')
}

async function onRefreshMails() {
    await fetchMails()
}

watch(
    () => [route.params.accountId, route.params.folder, route.query.is_active],
    () => {
        page.value = 1
        fetchMails()
    },
)

onMounted(fetchMails)
</script>

<template>
    <div class="page-container">
        <el-card shadow="never" class="premium-card">
            <template #header>
                <div class="header-row">
                    <div class="title">正在查看: {{ email }} - {{ folderText }}</div>
                    <div class="header-actions">
                        <el-button :icon="Refresh" :loading="loading" @click="onRefreshMails">刷新列表</el-button>
                        <el-button @click="goBack">返回上一页</el-button>
                    </div>
                </div>
            </template>

            <el-table v-loading="loading" :data="rows" border stripe>
                <el-table-column prop="from_name" label="发件人" min-width="220" show-overflow-tooltip>
                    <template #default="{ row }">
                        <span>{{ row.from_name || row.from_email || '-' }}</span>
                    </template>
                </el-table-column>
                <el-table-column prop="subject" label="主题" min-width="320" show-overflow-tooltip />
                <el-table-column prop="date" label="接收时间" min-width="180" />
                <el-table-column label="操作" width="120" fixed="right">
                    <template #default="{ row }">
                        <el-button type="primary" link @click="onViewDetail(row)">查看详情</el-button>
                    </template>
                </el-table-column>
            </el-table>

            <div class="pager">
                <el-pagination v-model:current-page="page" v-model:page-size="pageSize"
                    layout="total, sizes, prev, pager, next, jumper" :total="total" :page-count="pageCount"
                    :page-sizes="[10, 20, 50]" @change="fetchMails" />
            </div>
        </el-card>

        <el-dialog v-model="detailDialogVisible" width="960px" title="邮件详情" top="4vh" class="mail-detail-dialog">
            <el-skeleton v-if="detailLoading" :rows="8" animated />
            <template v-else>
                <template v-if="currentDetail">
                    <div class="detail-shell">
                        <div class="detail-meta">
                            <div><b>主题: </b>{{ currentDetail.subject || '-' }}</div>
                            <div><b>发件人: </b>{{ currentDetail.from || '-' }}</div>
                            <div><b>收件人: </b>{{ currentDetail.to || '-' }}</div>
                            <div><b>时间: </b>{{ currentDetail.date || '-' }}</div>
                        </div>

                        <div class="mail-body-panel">
                            <div class="mail-body-toolbar">正文预览</div>
                            <iframe
                                v-if="renderedBodyHtml"
                                class="mail-body-frame"
                                :srcdoc="renderedBodySrcdoc"
                                sandbox="allow-popups allow-popups-to-escape-sandbox"
                                referrerpolicy="no-referrer"
                            />
                            <pre v-else class="mail-text">{{ currentDetail.body_text || '(无正文)' }}</pre>
                        </div>
                    </div>
                </template>
                <el-empty v-else description="暂无详情" />
            </template>
        </el-dialog>
    </div>
</template>

<style scoped>
.header-row {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 12px;
}

.title {
    font-size: 16px;
    font-weight: 600;
}

.header-actions {
    display: flex;
    align-items: center;
    gap: 8px;
}

.pager {
    margin-top: 16px;
    display: flex;
    justify-content: flex-end;
}

.detail-shell {
    display: grid;
    gap: 18px;
}

.detail-meta {
    display: grid;
    gap: 8px;
    padding: 2px 2px 0;
    color: #606266;
    line-height: 1.7;
    word-break: break-word;
}

.mail-body-panel {
    overflow: hidden;
    border: 1px solid #dbe4ee;
    border-radius: 18px;
    background: linear-gradient(180deg, #ffffff, #f8fbff);
    box-shadow: 0 16px 34px rgba(15, 23, 42, 0.08);
}

.mail-body-toolbar {
    padding: 12px 16px;
    border-bottom: 1px solid #e2e8f0;
    background: linear-gradient(90deg, #eff6ff, #f8fafc);
    font-size: 13px;
    font-weight: 700;
    color: #475569;
    letter-spacing: 0.04em;
}

.mail-body-frame {
    display: block;
    width: 100%;
    min-height: 56vh;
    border: none;
    background: #fff;
}

.mail-text {
    margin: 0;
    padding: 18px;
    min-height: 62vh;
    background: #fff;
    color: #1f2937;
    line-height: 1.7;
    white-space: pre-wrap;
}

:deep(.mail-detail-dialog .el-dialog) {
    border-radius: 20px;
    overflow: hidden;
}

:deep(.mail-detail-dialog .el-dialog__header) {
    padding: 18px 22px 14px;
    border-bottom: 1px solid #e2e8f0;
    background: linear-gradient(180deg, #ffffff, #f8fafc);
}

:deep(.mail-detail-dialog .el-dialog__body) {
    padding: 20px 22px 22px;
    background: linear-gradient(180deg, #f8fafc, #f1f5f9);
}
</style>
