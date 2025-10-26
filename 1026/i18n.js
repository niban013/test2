// 🌐 i18n.js - 多國語言模組（支援位置參數）
(function (global) {
  const i18n = {
    data: {},
    currentLang: localStorage.getItem("lang") || "zh",
    ready: false,

    async init(lang) {
      await this.load(lang || this.currentLang);
    },

    async load(lang) {
      const url = `lang/${lang}.json`;
      try {
        const response = await fetch(url);
        if (!response.ok) throw new Error(`語言包載入失敗: ${url}`);
        this.data = await response.json();
        this.currentLang = lang;
        localStorage.setItem("lang", lang);
        this.ready = true;
      } catch (err) {
        console.error("❌ i18n 載入錯誤:", err);
      }
    },

    /**
     * 取得翻譯文字，支援多層 key（a.b.c）
     * 並支援位置參數 {0}, {1}, ...
     * i18n.t("welcome", "王小明") → "你好，王小明！"
     */
    t(key, ...args) {
      const keys = key.split(".");
      let val = this.data;
      for (const k of keys) {
        val = val?.[k];
        if (val === undefined) break;
      }
      if (val === undefined || val === null) return key;

      // 如果是字串且有位置參數
      if (typeof val === "string" && args.length > 0) {
        val = val.replace(/\{(\d+)\}/g, (match, index) => {
          return args[index] !== undefined ? args[index] : match;
        });
      }
      return val;
    },

    translateText(text) {
      // 找 menuItems 對應
      const map = this.data?.menuItems || {};
      return map[text.trim()] || text;
    },

    apply(target) {
      if (!this.ready) return;
      const $target = $(target || document);
      $target.find("[data-i18n]").each((_, el) => {
        const $el = $(el);
        const key = $el.attr("data-i18n");
        const argsAttr = $el.attr("data-i18n-args"); // 可選，逗號分隔
        let args = [];
        if (argsAttr) {
          args = argsAttr.split(",").map(s => s.trim());
        }
        const value = this.t(key, ...args);

        if ($el.is("input,button")) {
          $el.val(value).text(value);
        } else {
          $el.html(value);
        }
      });
    },

    async setLang(lang) {
      await this.load(lang);
      this.apply(document);
    },

    getLang() {
      return this.currentLang;
    }
  };

  global.i18n = i18n;
})(window);
