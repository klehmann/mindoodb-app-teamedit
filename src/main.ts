import { createApp } from "vue";
import PrimeVue from "primevue/config";
import "primeicons/primeicons.css";
import "@milkdown/crepe/theme/common/style.css";
import "@milkdown/crepe/theme/frame.css";

import App from "./App.vue";
import "@/assets/styles/main.css";
import { applyAppTheme, buildPrimeVueTheme } from "@/lib/theme";

const app = createApp(App);

app.use(PrimeVue, {
  ripple: true,
  theme: buildPrimeVueTheme(),
});

applyAppTheme();
app.mount("#app");
