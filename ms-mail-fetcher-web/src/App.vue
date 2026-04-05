<script setup>
import { computed, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import { Fold, Expand, User, Box } from '@element-plus/icons-vue'
import { getAccounts, getUiPreferences, updateUiPreferences } from '@/api/accounts'

const SIDEBAR_COLLAPSE_STORAGE_KEY = 'ms_mail_fetcher_sidebar_collapsed'

function readSidebarCollapsedStateFromLocal() {
  try {
    return localStorage.getItem(SIDEBAR_COLLAPSE_STORAGE_KEY) === '1'
  } catch {
    return false
  }
}

function writeSidebarCollapsedStateToLocal(isCollapsed) {
  try {
    localStorage.setItem(SIDEBAR_COLLAPSE_STORAGE_KEY, isCollapsed ? '1' : '0')
  } catch {
    // ignore storage failures
  }
}

const isCollapse = ref(false)
const preferenceLoaded = ref(false)
const activeCount = ref(0)
const archivedCount = ref(0)

async function loadSidebarCollapsedState() {
  try {
    const data = await getUiPreferences()
    if (typeof data?.sidebar_collapsed === 'boolean') {
      isCollapse.value = data.sidebar_collapsed
      writeSidebarCollapsedStateToLocal(data.sidebar_collapsed)
      return
    }
  } catch {
    // ignore backend failures, fallback to local
  }

  isCollapse.value = readSidebarCollapsedStateFromLocal()
}

const toggleCollapse = () => {
  isCollapse.value = !isCollapse.value
}

watch(isCollapse, async (value) => {
  if (!preferenceLoaded.value) {
    return
  }

  writeSidebarCollapsedStateToLocal(value)

  try {
    await updateUiPreferences({ sidebar_collapsed: value })
  } catch {
    // ignore backend failures
  }
})

const totalCount = computed(() => activeCount.value + archivedCount.value)

async function fetchOverviewCounts() {
  try {
    const [activeData, archivedData] = await Promise.all([
      getAccounts({ is_active: true, page: 1, page_size: 1 }),
      getAccounts({ is_active: false, page: 1, page_size: 1 }),
    ])
    activeCount.value = Number(activeData?.total || 0)
    archivedCount.value = Number(archivedData?.total || 0)
  } catch {
    // ignore stats errors in header
  }
}

function onAccountsChanged() {
  fetchOverviewCounts()
}


onMounted(async () => {
  await loadSidebarCollapsedState()
  preferenceLoaded.value = true
  fetchOverviewCounts()
  window.addEventListener('accounts-changed', onAccountsChanged)
})

onBeforeUnmount(() => {
  window.removeEventListener('accounts-changed', onAccountsChanged)
})
</script>

<template>
  <el-container class="layout-root">
    <el-aside
      :width="isCollapse ? '64px' : '210px'"
      class="aside"
      style="transition: width 0.3s; overflow: hidden"
    >
      <div class="brand">
        <span v-if="!isCollapse">MS-Mail GPT Manager</span>
        <span v-else>MS</span>
      </div>
      <el-menu :default-active="$route.path" router :collapse="isCollapse" class="el-menu-vertical">
        <el-menu-item index="/active">
          <el-icon>
            <User />
          </el-icon>
          <template #title>活跃账号池</template>
        </el-menu-item>
        <el-menu-item index="/archived">
          <el-icon>
            <Box />
          </el-icon>
          <template #title>已归档账号</template>
        </el-menu-item>
      </el-menu>
    </el-aside>

    <el-container class="right-container">
      <el-header class="header">
        <div class="header-left">
          <el-icon class="collapse-btn" @click="toggleCollapse">
            <component :is="isCollapse ? Expand : Fold" />
          </el-icon>
          <span class="title">微软邮箱账号管理系统</span>
        </div>

        <div class="overview">
          <div class="overview-pill active-pill">
            <div class="pill-dot" style="background: #10b981"></div>
            <span
              >活跃: <strong>{{ activeCount }}</strong></span
            >
          </div>
          <div class="overview-pill archived-pill">
            <div class="pill-dot" style="background: #6b7280"></div>
            <span
              >归档: <strong>{{ archivedCount }}</strong></span
            >
          </div>
          <div class="overview-pill total-pill">
            <div class="pill-dot" style="background: #4f46e5"></div>
            <span
              >总计: <strong>{{ totalCount }}</strong></span
            >
          </div>
        </div>
      </el-header>
      <el-main class="main">
        <router-view v-slot="{ Component }">
          <transition name="fade" mode="out-in">
            <component :is="Component" />
          </transition>
        </router-view>
      </el-main>
    </el-container>
  </el-container>
</template>

<style scoped>
.layout-root {
  height: 100vh;
  overflow: hidden;
  background: #f5f7fa;
}

.aside {
  background: #fff;
  border-right: 1px solid #ebeef5;
  display: flex;
  flex-direction: column;
}

.el-menu-vertical {
  border-right: none;
}

.brand {
  font-size: 15px;
  font-weight: 700;
  height: 60px;
  display: flex;
  align-items: center;
  justify-content: center;
  border-bottom: 1px solid #ebeef5;
  white-space: nowrap;
  overflow: hidden;
}

.right-container {
  display: flex;
  flex-direction: column;
}

.header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  background: rgba(255, 255, 255, 0.85);
  backdrop-filter: blur(12px);
  border-bottom: 1px solid #f1f5f9;
  height: 64px;
  padding: 0 24px;
  box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.03);
}

.header-left {
  display: flex;
  align-items: center;
  gap: 16px;
}

.collapse-btn {
  font-size: 20px;
  cursor: pointer;
  color: #64748b;
  transition: color 0.3s;
}

.collapse-btn:hover {
  color: var(--el-color-primary);
}

.title {
  font-size: 17px;
  font-weight: 600;
  color: #0f172a;
  letter-spacing: -0.01em;
}

.overview {
  display: flex;
  align-items: center;
  gap: 12px;
}

.overview-pill {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 6px 14px;
  border-radius: 20px;
  font-size: 13px;
  font-weight: 500;
  background: #ffffff;
  border: 1px solid #e2e8f0;
  color: #475569;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.02);
  transition: all 0.2s ease;
  border-radius: 0%;
}

.overview-pill:hover {
  transform: translateY(-1px);
  box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
}

.active-pill {
  background: #ecfdf5;
  border-color: #d1fae5;
  color: #065f46;
}

.archived-pill {
  background: #f3f4f6;
  border-color: #e5e7eb;
  color: #374151;
}

.total-pill {
  background: #eef2ff;
  border-color: #e0e7ff;
  color: #3730a3;
}

.pill-dot {
  width: 8px;
  height: 8px;
  border-radius: 50%;
}

.main {
  padding: 24px;
  overflow-x: hidden;
  overflow-y: auto;
  display: flex;
  flex-direction: column;
  flex: 1;
}
</style>


