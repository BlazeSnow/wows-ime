import { defineConfig } from 'vitepress'
import { shared } from './shared'
import { zh } from './zh'
import { zhHant } from './zh-hant'
import { ja } from './ja'

export default defineConfig({
    ...shared,
    locales: {
        root: { label: '简体中文', ...zh },
        'zh-hant': { label: '繁體中文', link: '/zh-hant/', ...zhHant },
        ja: { label: '日本語', link: '/ja/', ...ja }
    }
})
