import { defineConfig } from 'vitepress'

//日本語（/ja/ 路径）语言配置
export const ja = defineConfig({
    //网页语言
    lang: 'ja-JP',
    //网页描述
    description: 'World of Warships の日本語入力設定ファイルを編集し、より多くの入力メソッドを使えるようにするツールです。',
    //主题配置
    themeConfig: {
        //语言切换按钮提示
        langMenuLabel: '言語を切り替える',
        //切换深色或浅色模式提示
        darkModeSwitchLabel: 'ダークモードとライトモードを切り替える',
        //切换至浅色模式提示
        lightModeSwitchTitle: 'ライトモードに切り替える',
        //切换至深色模式提示
        darkModeSwitchTitle: 'ダークモードに切り替える',
        //目录按钮文字
        sidebarMenuLabel: 'メニュー',
        //回到顶部文字
        returnToTopLabel: 'ページの先頭へ',
        //右边的小目录标题
        outlineTitle: 'このページの内容',
        //上一篇下一篇
        docFooter: {
            prev: '前へ',
            next: '次へ'
        },
        nav: [
            {
                text: 'ホーム',
                link: '/ja/'
            },
            {
                text: '更新履歴',
                link: '/ja/CHANGELOG',
                activeMatch: '/ja/CHANGELOG'
            }
        ],
        notFound: {
            title: 'ページが見つかりません',
            quote: '申し訳ありません。お探しのページが見つかりませんでした',
            linkLabel: 'ホームへ戻る',
            linkText: 'ホームへ戻る',
            code: '404',
        },
        //主页页脚
        footer: {
            message: 'このソフトウェアは <a href="https://www.gnu.org/licenses/agpl-3.0.html" target="_blank">GNU AGPL v3.0</a> ライセンスで提供されています。',
            copyright: 'Copyright © 2026 <a href="https://github.com/BlazeSnow" target="_blank">BlazeSnow</a>.'
        }
    }
})
