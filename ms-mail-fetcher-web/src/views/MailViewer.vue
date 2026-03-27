<script setup>
import { computed, onMounted, ref, watch } from 'vue'
import { useRoute, useRouter } from 'vue-router'
import { ElMessage } from 'element-plus'
import { getMailDetail, getMailList } from '@/api/accounts'

const route = useRoute()
const router = useRouter()

const accountId = computed(() => Number(route.params.accountId))
const folder = computed(() => String(route.params.folder || 'inbox'))
const email = computed(() => String(route.query.email || `账号 #${accountId.value}`))

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

async function fetchMails() {
    if (!accountId.value) return

    loading.value = true
    try {
        const data = await getMailList(accountId.value, folder.value, {
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
        const data = await getMailDetail(accountId.value, folder.value, row.uid)
        currentDetail.value = data.detail
    } catch (error) {
        currentDetail.value = null
        ElMessage.error(error.message || '读取邮件详情失败')
    } finally {
        detailLoading.value = false
    }
}

function goBack() {
    if (window.history.length > 1) {
        router.back()
    } else {
        router.push('/active')
    }
}

watch(
    () => [route.params.accountId, route.params.folder],
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
                <el-button @click="goBack">返回上一页</el-button>
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

    <el-dialog v-model="detailDialogVisible" width="900px" title="邮件详情" top="4vh">
        <el-skeleton v-if="detailLoading" :rows="8" animated />
        <template v-else>
            <template v-if="currentDetail">
                <div class="detail-meta">
                    <div><b>主题：</b>{{ currentDetail.subject || '-' }}</div>
                    <div><b>发件人：</b>{{ currentDetail.from || '-' }}</div>
                    <div><b>收件人：</b>{{ currentDetail.to || '-' }}</div>
                    <div><b>时间：</b>{{ currentDetail.date || '-' }}</div>
                </div>
                <el-divider />
                <div class="mail-body" v-if="renderedBodyHtml" v-html="renderedBodyHtml"></div>
                <pre v-else class="mail-text">{{ currentDetail.body_text || '(无正文)' }}</pre>
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

.pager {
    margin-top: 16px;
    display: flex;
    justify-content: flex-end;
}

.detail-meta {
    display: grid;
    gap: 8px;
    color: #606266;
}

.mail-body {
    max-height: 60vh;
    overflow: auto;
}

.mail-text {
    margin: 0;
    white-space: pre-wrap;
}
</style>
