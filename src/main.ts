import { createApp } from "vue";
import PrimeVue from "primevue/config";
import "@fontsource/inter/latin-400.css";
import "@fontsource/inter/latin-600.css";
import "@fontsource/inter/latin-700.css";
import "@fontsource/rubik/latin-400.css";
import "@fontsource/rubik/latin-700.css";
import "@fontsource/jetbrains-mono/latin-400.css";
import "primeicons/primeicons.css";
import "@milkdown/crepe/theme/common/style.css";

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
