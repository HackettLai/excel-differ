// copyright.js

class Copyright {
    constructor() {
        this.year = new Date().getFullYear();
        this.author = 'Hackett.Lai';
        this.github = 'https://github.com/Hackett0/Excel-Differ';
    }

    init() {
        // 更新年份
        const yearElement = document.getElementById('thisYear');
        if (yearElement) {
            yearElement.textContent = this.year;
        }

        // 更新完整的 copyright 文字（如果有這個 element）
        const copyrightElement = document.getElementById('copyrightText');
        if (copyrightElement) {
            copyrightElement.innerHTML = `
                ${this.author} @ ${this.year} | 
                <a href="${this.github}" target="_blank" style="color: white; text-decoration: underline;">
                    View it on github
                </a>
            `;
        }
    }
}

export default Copyright;  // ✅ 一定要有這行！