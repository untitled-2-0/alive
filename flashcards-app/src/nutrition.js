// Nutrition: nutrient table, reference values, day totals, weekly shortfall analysis.
// Pure logic + the food catalogue loader. Storage and UI live in FlashcardsApp.jsx.
//
// Every number here is a reference approximation. Food composition varies by variety,
// season, soil and cooking; a tracker total is an orientation, never a diagnosis.

/* ---------- nutrient table ---------- */
// kind: "goal"    — aim to reach it (bar fills up, low is the problem)
//       "ceiling" — aim to stay under it (low is fine, high is the problem)
export const NUTRIENTS = [
  { k: "kcal",       label: "Калорії",     unit: "ккал", cat: "energy", kind: "goal", info: "Енергія на все: дихання, роботу мозку, тепло, рух. З'їла більше, ніж витратила — тіло відкладає запас; менше — бере зі своїх." },
  { k: "protein",    label: "Білки",       unit: "г",    cat: "macro",  kind: "goal", info: "Будівельний матеріал: м'язи, шкіра, волосся, нігті, ферменти, антитіла. У дефіциті калорій саме білок тримає м'язи, щоб худнути жиром, а не ними." },
  { k: "fat",        label: "Жири",        unit: "г",    cat: "macro",  kind: "goal", info: "Потрібен для гормонів (зокрема статевих), оболонок клітин і засвоєння вітамінів A, D, E, K. Надто мало жиру — сухість шкіри й збої циклу." },
  { k: "satFat",     label: "з них насичені", unit: "г", cat: "macro",  kind: "ceiling", sub: true, info: "Насичені жири. Трохи потрібно, але надлишок піднімає «поганий» холестерин і навантажує судини. Тому це стеля, а не ціль." },
  { k: "carbs",      label: "Вуглеводи",   unit: "г",    cat: "macro",  kind: "goal", info: "Головне швидке паливо для мозку й м'язів — мозок працює майже виключно на глюкозі. Найкращі джерела — крупи, овочі, фрукти, бобові." },
  { k: "sugar",      label: "з них цукри", unit: "г",    cat: "macro",  kind: "ceiling", sub: true, info: "Вуглеводи, що засвоюються миттєво: різкий стрибок енергії й таке саме падіння, плюс навантаження на зуби й обмін. Стеля, а не ціль." },
  { k: "fiber",      label: "Клітковина",  unit: "г",    cat: "macro",  kind: "goal", info: "Клітковина не засвоюється, але годує мікрофлору кишківника, дає ситість, вирівнює цукор у крові й налагоджує травлення." },

  { k: "iron",       label: "Залізо",      unit: "мг",   cat: "min", kind: "goal", info: "Переносить кисень по крові у складі гемоглобіну. Нестача — втома, задишка від дрібних навантажень, випадіння волосся. У жінок витрачається з місячними, тому норма вища. Вітамін C у тій самій страві помітно покращує засвоєння заліза з рослинної їжі." },
  { k: "zinc",       label: "Цинк",        unit: "мг",   cat: "min", kind: "goal", info: "Імунітет, загоєння ран, шкіра й волосся, відчуття смаку та нюху, гормональний баланс." },
  { k: "copper",     label: "Мідь",        unit: "мг",   cat: "min", kind: "goal", info: "Допомагає тілу засвоювати залізо, бере участь у творенні сполучної тканини, судин і пігменту волосся." },
  { k: "magnesium",  label: "Магній",      unit: "мг",   cat: "min", kind: "goal", info: "Задіяний у сотнях реакцій: розслаблення м'язів, робота нервів, сон, серцевий ритм. Нестача часто відчувається як судоми в литках, тривожність і важке засинання." },
  { k: "calcium",    label: "Кальцій",     unit: "мг",   cat: "min", kind: "goal", info: "Кістки й зуби, а ще скорочення м'язів і передача нервових сигналів. Без вітаміну D засвоюється погано — вони працюють у парі." },
  { k: "potassium",  label: "Калій",       unit: "мг",   cat: "min", kind: "goal", info: "Тримає водний баланс і тиск, потрібен для роботи серця й м'язів. Працює в парі з натрієм: більше калію — м'якший вплив солі на тиск." },
  { k: "phosphorus", label: "Фосфор",      unit: "мг",   cat: "min", kind: "goal", info: "Разом із кальцієм будує кістки й входить до молекули енергії (АТФ), якою користується кожна клітина." },
  { k: "selenium",   label: "Селен",       unit: "мкг",  cat: "min", kind: "goal", info: "Антиоксидант — захищає клітини від пошкодження. Потрібен щитоподібній залозі для роботи з гормонами." },
  { k: "iodine",     label: "Йод",         unit: "мкг",  cat: "min", kind: "goal", info: "Сировина для гормонів щитоподібної залози, які керують обміном речовин, температурою тіла й рівнем енергії." },
  { k: "sodium",     label: "Натрій (сіль)", unit: "мг", cat: "min", kind: "ceiling", info: "Натрій — це сіль. Потрібен нервам і водному балансу, але надлишок піднімає тиск і затримує воду. Тому тут стеля, а не ціль." },

  { k: "vitA",       label: "Вітамін A",   unit: "мкг",  cat: "vit", kind: "goal", info: "Зір, особливо в сутінках, а також шкіра, слизові оболонки й імунітет. Жиророзчинний — засвоюється разом із жиром." },
  { k: "vitC",       label: "Вітамін C",   unit: "мг",   cat: "vit", kind: "goal", info: "Імунітет і синтез колагену — це шкіра, судини, загоєння. Ще й різко покращує засвоєння заліза з рослинної їжі, якщо їсти їх разом." },
  { k: "vitD",       label: "Вітамін D",   unit: "мкг",  cat: "vit", kind: "goal", info: "Без нього кальцій майже не засвоюється. Впливає на кістки, імунітет і настрій. Утворюється в шкірі від сонця, тому взимку його бракує найчастіше." },
  { k: "vitE",       label: "Вітамін E",   unit: "мг",   cat: "vit", kind: "goal", info: "Антиоксидант: захищає оболонки клітин і шкіру від пошкодження." },
  { k: "vitK",       label: "Вітамін K",   unit: "мкг",  cat: "vit", kind: "goal", info: "Потрібен для згортання крові — щоб ранки загоювались, і для того, щоб кальцій ішов у кістки." },
  { k: "b1",         label: "B1 тіамін",   unit: "мг",   cat: "vit", kind: "goal", info: "Тіамін. Перетворює вуглеводи на енергію й підтримує роботу нервової системи." },
  { k: "b2",         label: "B2 рибофлавін", unit: "мг", cat: "vit", kind: "goal", info: "Рибофлавін. Обмін енергії, здоров'я шкіри, слизових і очей. Заїди в куточках рота — класична ознака нестачі." },
  { k: "b3",         label: "B3 ніацин",   unit: "мг",   cat: "vit", kind: "goal", info: "Ніацин. Бере участь в енергетичному обміні кожної клітини, потрібен шкірі й нервовій системі." },
  { k: "b5",         label: "B5 пантотенова", unit: "мг", cat: "vit", kind: "goal", info: "Пантотенова кислота. Потрібна, щоб з їжі діставалась енергія — бере участь у переробці жирів, білків і вуглеводів. Є майже в усьому потроху, тому справжня нестача рідкісна." },
  { k: "b6",         label: "B6",          unit: "мг",   cat: "vit", kind: "goal", info: "Обмін білка, творення гемоглобіну і нейромедіаторів — зокрема серотоніну. Тобто впливає на настрій і сон." },
  { k: "b7",         label: "B7 біотин",   unit: "мкг",  cat: "vit", kind: "goal", info: "Біотин. Обмін жирів і цукру, а ще стан волосся, нігтів і шкіри — саме тому його часто п'ють від випадіння волосся. Тіло частково отримує його від мікрофлори кишківника." },
  { k: "b9",         label: "B9 фолат",    unit: "мкг",  cat: "vit", kind: "goal", info: "Фолат. Потрібен для поділу клітин і творення крові. Критично важливий до зачаття й у першому триместрі вагітності." },
  { k: "b12",        label: "B12",         unit: "мкг",  cat: "vit", kind: "goal", info: "Нерви, творення еритроцитів, енергія. Є практично лише у тваринній їжі, тому на веганстві за ним стежать окремо." },

  // kind "info" — тіло синтезує їх саме, офіційної добової норми не існує.
  // Тому показуємо тільки скільки з'їдено, без відсотків і без «мало / забагато».
  { k: "taurine",  label: "Таурин",   unit: "мг", cat: "other", kind: "info", info: "Амінокислота, яку тіло виробляє саме. Задіяна в роботі серця, м'язів, зору й нервової системи. З їжі надходить тільки з тваринних продуктів — найбільше з морепродуктів і темного м'яса; у рослинній їжі її практично нема. Офіційної добової норми не існує, тому тут лише кількість, без відсотків." },
  { k: "inositol", label: "Інозитол", unit: "мг", cat: "other", kind: "info", info: "Його часто називають «вітаміном B8», але це не вітамін: тіло синтезує близько грама на день саме. Бере участь у передачі сигналів усередині клітин і в обміні жирів. Найбільше в висівках, бобових, цитрусових і дині. Офіційної добової норми немає, тому тут лише кількість.\n\nОкремо про добавки — це опис досліджень, а не призначення. У дослідженнях інозитол давали грамами, тобто в тисячі разів більше, ніж дає їжа: з тарілки таких кількостей не набрати. Найкраще доведена дія при СПКЯ та інсулінорезистентності — 4 г міо-інозитолу на день (зазвичай по 2 г двічі), часто в парі з D-хіро-інозитолом у співвідношенні 40:1; покращує чутливість до інсуліну, овуляцію та регулярність циклу. При панічному розладі вивчали 12–18 г/день, при депресії — близько 12 г/день, при ОКР — 18 г/день, але тут досліджень мало й результати суперечливі, тож вважати це доведеним не можна. Переноситься зазвичай добре; від ~12 г бувають нудота, гази й розлад шлунка. При ендокринних діагнозах, вагітності чи прийомі ліків дозу має підбирати лікар — інозитол впливає на цукор крові і може складатись із дією інших препаратів." },
];

export const CATS = [
  { id: "macro", label: "Макронутрієнти" },
  { id: "min",   label: "Мінерали" },
  { id: "vit",   label: "Вітаміни" },
  { id: "other", label: "Інше" },
];

export const NUTRIENT_KEYS = NUTRIENTS.map((n) => n.k);
export const NUTRIENT_BY_KEY = Object.fromEntries(NUTRIENTS.map((n) => [n.k, n]));

/* ---------- reference daily values ---------- */
// Adult RDA/AI (EFSA + US DRI, rounded). Sex-specific where it genuinely differs.
// Upper limits (`ul`) are the levels worth flagging, not hard danger thresholds.
const UL = {
  iron: 45, zinc: 40, copper: 10, magnesium: 350, calcium: 2500, selenium: 400,
  iodine: 1100, sodium: 2300, vitA: 3000, vitD: 100, vitE: 1000, vitC: 2000,
  b3: 35, b6: 100, b9: 1000,
};

export function referenceValues({ sex = "f", age = 30, weightKg = 65, kcal = 2000 } = {}) {
  const f = sex !== "m";
  const older = age >= 51;
  const w = weightKg > 0 ? weightKg : 65;
  const r = {
    kcal,
    // 1.6 г/кг — вище офіційного мінімуму 0.8, бо в дефіциті білок зберігає м'язи
    protein: Math.round(1.6 * w),
    fat: Math.round((kcal * 0.3) / 9),
    carbs: Math.round((kcal * 0.45) / 4),
    satFat: Math.round((kcal * 0.1) / 9),          // стеля
    sugar: Math.round((kcal * 0.1) / 4),           // стеля
    fiber: f ? 25 : 38,

    iron: f && !older ? 18 : 8,
    zinc: f ? 8 : 11,
    copper: 0.9,
    magnesium: f ? 320 : 420,
    calcium: older ? 1200 : 1000,
    potassium: f ? 2600 : 3400,
    phosphorus: 700,
    selenium: 55,
    iodine: 150,
    sodium: 2300,                                   // стеля

    vitA: f ? 700 : 900,
    vitC: f ? 75 : 90,
    vitD: age >= 71 ? 20 : 15,
    vitE: 15,
    vitK: f ? 90 : 120,
    b1: f ? 1.1 : 1.2,
    b2: f ? 1.1 : 1.3,
    b3: f ? 14 : 16,
    b6: older ? (f ? 1.5 : 1.7) : 1.3,
    b5: 5,
    b7: 30,
    b9: 400,
    b12: 2.4,
  };
  return r;
}

export const upperLimit = (k) => UL[k];

/* ---------- day maths ---------- */
export function emptyTotals() {
  const t = {};
  for (const k of NUTRIENT_KEYS) t[k] = 0;
  return t;
}

// entries: [{ food, grams }] — food is a full per-100 g record
export function sumEntries(entries) {
  const t = emptyTotals();
  for (const e of entries || []) {
    if (!e || !e.food) continue;
    const f = (Number(e.grams) || 0) / 100;
    for (const k of NUTRIENT_KEYS) {
      const v = Number(e.food[k]);
      if (isFinite(v)) t[k] += v * f;
    }
  }
  for (const k of NUTRIENT_KEYS) t[k] = Math.round(t[k] * 100) / 100;
  return t;
}

export function pct(value, target) {
  if (!target || !isFinite(target)) return null;
  return Math.round((value / target) * 100);
}

// Status for one nutrient on one day.
export function statusOf(k, value, refs) {
  const spec = NUTRIENT_BY_KEY[k];
  if (spec && spec.kind === "info") return { p: null, level: "info" };
  const target = refs[k];
  const p = pct(value, target);
  const ul = UL[k];
  if (p == null) return { p: null, level: "none" };
  if (spec && spec.kind === "ceiling") {
    if (p > 130) return { p, level: "over" };
    if (p > 100) return { p, level: "high" };
    return { p, level: "ok" };
  }
  if (ul && value > ul) return { p, level: "over" };
  if (p < 50) return { p, level: "low" };
  if (p < 90) return { p, level: "under" };
  return { p, level: "ok" };
}

/* ---------- weekly analysis ---------- */
// days: [{ date, totals }] — newest last. Returns what is chronically low / high.
export function weekAnalysis(days, refs) {
  const usable = (days || []).filter((d) => d && d.totals && d.totals.kcal > 0);
  const short = [];
  const excess = [];
  if (usable.length < 2) return { short, excess, daysCounted: usable.length };

  for (const n of NUTRIENTS) {
    if (n.k === "kcal" || n.kind === "info") continue;
    const target = refs[n.k];
    if (!target) continue;
    const vals = usable.map((d) => Number(d.totals[n.k]) || 0);
    const avg = vals.reduce((a, b) => a + b, 0) / vals.length;
    const avgPct = Math.round((avg / target) * 100);
    if (n.kind === "ceiling") {
      const daysOver = vals.filter((v) => v > target).length;
      if (avgPct > 110 && daysOver >= Math.ceil(usable.length / 2)) {
        excess.push({ k: n.k, label: n.label, unit: n.unit, avg: round1(avg), target, avgPct, days: daysOver });
      }
      continue;
    }
    const ul = UL[n.k];
    const daysOverUL = ul ? vals.filter((v) => v > ul).length : 0;
    if (daysOverUL >= 2) {
      excess.push({ k: n.k, label: n.label, unit: n.unit, avg: round1(avg), target: ul, avgPct, days: daysOverUL, isUL: true });
      continue;
    }
    const daysLow = vals.filter((v) => v < target * 0.7).length;
    if (avgPct < 70 && daysLow >= Math.ceil(usable.length / 2)) {
      short.push({ k: n.k, label: n.label, unit: n.unit, avg: round1(avg), target, avgPct, days: daysLow });
    }
  }
  short.sort((a, b) => a.avgPct - b.avgPct);
  excess.sort((a, b) => b.avgPct - a.avgPct);
  return { short, excess, daysCounted: usable.length };
}

const round1 = (v) => Math.round(v * 10) / 10;

// Best sources for a nutrient, out of the catalogue. Skips foods nobody eats 100 g of.
const SKIP_AS_SOURCE = new Set(["salt", "sugar", "oilSunflower", "oilOlive", "sunflowerOil", "oliveOil"]);
export function bestSources(foods, key, limit = 5) {
  return (foods || [])
    .filter((f) => isFinite(Number(f[key])) && Number(f[key]) > 0 && !SKIP_AS_SOURCE.has(f.id))
    .sort((a, b) => Number(b[key]) - Number(a[key]))
    .slice(0, limit)
    .map((f) => ({ id: f.id, name: f.name, amount: round1(Number(f[key])) }));
}

/* ---------- own dishes ----------
   Страва — це список інгредієнтів у грамах. Її склад рахується як сума інгредієнтів,
   поділена на загальну вагу, тому далі вона поводиться як звичайний продукт:
   є значення на 100 г, а «вся страва» — просто порція розміром у цю вагу.
   Вагу можна задати вручну (напр. після варіння вода випарувалась). */
export function composeDish({ id, name, items, cookedGrams }) {
  const list = (items || []).filter((i) => i && i.food && Number(i.grams) > 0);
  const rawGrams = list.reduce((s, i) => s + Number(i.grams), 0);
  const totalGrams = Number(cookedGrams) > 0 ? Number(cookedGrams) : rawGrams;
  const sum = sumEntries(list);
  const dish = {
    id: id || "dish_" + Math.random().toString(36).slice(2, 9),
    name: (name || "Моя страва").trim(),
    group: "Мої страви",
    dish: true,
    items: list.map((i) => ({ name: i.food.name, grams: Number(i.grams), food: pickNutrients(i.food) })),
    rawGrams,
    portion: Math.round(totalGrams),
    portionLabel: "вся страва",
  };
  const k = totalGrams > 0 ? 100 / totalGrams : 0;
  for (const key of NUTRIENT_KEYS) dish[key] = Math.round((sum[key] || 0) * k * 1000) / 1000;
  return dish;
}

export function pickNutrients(food) {
  const out = {};
  for (const k of NUTRIENT_KEYS) out[k] = Number(food[k]) || 0;
  return out;
}

/* ---------- catalogue ---------- */
let CATALOGUE = null;
export async function loadFoods() {
  if (CATALOGUE) return CATALOGUE;
  try {
    const res = await fetch("/foods.json");
    if (!res.ok) throw new Error(`foods.json ${res.status}`);
    CATALOGUE = await res.json();
  } catch (e) {
    console.error("[nutrition] catalogue unavailable:", e);
    CATALOGUE = [];
  }
  return CATALOGUE;
}

export function searchFoods(list, query) {
  const q = (query || "").trim().toLowerCase();
  if (!q) return [];
  const starts = [];
  const contains = [];
  for (const f of list || []) {
    const n = (f.name || "").toLowerCase();
    if (n.startsWith(q)) starts.push(f);
    else if (n.includes(q)) contains.push(f);
  }
  return [...starts, ...contains].slice(0, 25);
}

/* ---------- on-demand nutrient profile ---------- */
// Same shape as the Languages generator: try the proxy, then a direct browser call.
// Without either (plain localhost / static host) there is no key, so this fails and
// the caller falls back to manual entry — it never invents numbers silently.
const FOOD_PROMPT = (name) => `Ти — довідник складу продуктів. Дай склад продукту «${name}» на 100 г.

Поверни ЛИШЕ JSON без пояснень, без markdown, точно з такими ключами:
{"name":"назва українською","group":"категорія українською","kcal":0,"protein":0,"fat":0,"satFat":0,"carbs":0,"sugar":0,"fiber":0,"iron":0,"zinc":0,"copper":0,"magnesium":0,"calcium":0,"potassium":0,"sodium":0,"phosphorus":0,"selenium":0,"iodine":0,"vitA":0,"vitC":0,"vitD":0,"vitE":0,"vitK":0,"b1":0,"b2":0,"b3":0,"b5":0,"b6":0,"b7":0,"b9":0,"b12":0}

Одиниці: kcal — ккал; protein, fat, satFat, carbs, sugar, fiber — грами; iron, zinc, copper, magnesium, calcium, potassium, sodium, phosphorus, vitC, vitE, b1, b2, b3, b5, b6 — мг; selenium, iodine, vitA, vitD, vitK, b7, b9, b12 — мкг. Усе на 100 г. Кожне поле обов'язкове, став 0 там, де справді нуль.`;

function extractJson(text) {
  try { return JSON.parse(text); } catch (e) { /* try to find the object */ }
  const s = text.indexOf("{");
  const e = text.lastIndexOf("}");
  if (s >= 0 && e > s) return JSON.parse(text.slice(s, e + 1));
  throw new Error("no JSON in response");
}

export async function generateFood(name) {
  const payload = JSON.stringify({
    model: "claude-sonnet-4-6",
    max_tokens: 1000,
    messages: [{ role: "user", content: FOOD_PROMPT(name) }],
  });
  const attempts = [
    { url: "/api/anthropic/messages", headers: { "content-type": "application/json" } },
    {
      url: "https://api.anthropic.com/v1/messages",
      headers: {
        "content-type": "application/json",
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true",
      },
    },
  ];
  let lastError = null;
  for (const attempt of attempts) {
    try {
      const res = await fetch(attempt.url, { method: "POST", headers: attempt.headers, body: payload });
      if (!res.ok) throw new Error(`generation failed (${res.status})`);
      const data = await res.json();
      const text = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n").trim();
      const raw = extractJson(text);
      return sanitizeFood(raw, name);
    } catch (e) {
      lastError = e;
    }
  }
  throw lastError || new Error("generation unavailable");
}

// Never trust the generated object: coerce every field, clamp the impossible.
export function sanitizeFood(raw, fallbackName) {
  const out = {
    id: "ai_" + Math.random().toString(36).slice(2, 9),
    name: String(raw && raw.name ? raw.name : fallbackName || "продукт").slice(0, 60),
    group: String(raw && raw.group ? raw.group : "Згенеровані").slice(0, 40),
    ai: true,
  };
  for (const k of NUTRIENT_KEYS) {
    const v = Number(raw && raw[k]);
    out[k] = isFinite(v) && v >= 0 ? Math.round(v * 1000) / 1000 : 0;
  }
  if (out.satFat > out.fat) out.satFat = out.fat;
  if (out.sugar > out.carbs) out.sugar = out.carbs;
  if (out.protein + out.fat + out.carbs > 100) {
    const s = out.protein + out.fat + out.carbs;
    out.protein = Math.round((out.protein / s) * 10000) / 100;
    out.fat = Math.round((out.fat / s) * 10000) / 100;
    out.carbs = Math.round((out.carbs / s) * 10000) / 100;
  }
  const atwater = 4 * out.protein + 9 * out.fat + 4 * out.carbs;
  if (!out.kcal || out.kcal > atwater * 1.6 + 40) out.kcal = Math.round(atwater);
  return out;
}
