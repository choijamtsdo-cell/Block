# Хантайшир Эко Блок — AI борлуулалтын зөвлөх

Энэ нь GitHub Pages вебсайтад Cloudflare Workers AI ашиглан Mongolian AI чат нэмэх бэлэн код.

**Одоогийн вэбсайт өөрийн AI API key хадгалахгүй.** Cloudflare Worker өөрийн Workers AI binding-ээр AI model ажиллуулна.

## Cloudflare Dashboard дээр байрлуулах

1. https://dash.cloudflare.com руу нэвтэрнэ. Cloudflare account байхгүй бол үүсгэнэ.
2. **Workers & Pages → Create application → Worker** сонгоно. Зарим dashboard дээр **Select a template** гарч болно; энгийн Worker/Hello World үүсгээд нэрийг `khantaishir-ai-adviser` гэж өгнө.
3. **Edit Code** хэсгийг нээгээд [index.js](./index.js)-ийн бүтэн кодыг хуулж хуучныг солиод Save/Deploy дарна.
4. Worker-ийн **Settings → Bindings → Add binding → Workers AI** хэсэгт binding-ийн нэрийг яг **AI** гэж өгөөд хадгалж Deploy хийнэ. Wrangler ашиглавал энэ фолдер дахь `wrangler.toml` бэлэн.
5. Worker URL `https://khantaishir-ai-adviser.<account>.workers.dev` хэлбэртэй гарна. Түүний `/health` төгсгөл рүү орж `"ok":true` эсэхийг шалгана.
6. Үүссэн **workers.dev Worker URL**-ээ ChatGPT-д илгээнэ. Сайтын `index.html` дахь `khantaishir-ai-endpoint` мета тохиргоонд `https://...workers.dev/api/chat` гэж холбож өгнө. URL тохируулаагүй бол чат товч огт харагдахгүй.

## Инженерийн тохиргоо (сонголт)

Node.js болон Wrangler ашиглавал:

```bash
cd cloudflare-ai
npx wrangler login
npx wrangler deploy
```

## AI юу мэдэх вэ?

- 300,000₮/м³ нь **урьдчилсан жишиг үнэ**; 200 мм ба 300 мм блокийн хэмжээ, ширхэгийн харьцаа.
- Говь-Алтай аймгийн Тайшир сум дахь үйлдвэр, жилд 50,000 м³ хүчин чадал.
- Блокийн суурь хэрэглээ, ерөнхий давуу тал, үнийн хүсэлт болон тооцоолуур руу хөтлөх.
- Тээвэр, хүргэлт, албан үнэ, НӨАТ, бэлэн нөөц, инженерийн баталгааг дур мэдэн амлахгүй.

Чат AI учраас алдаа гаргаж болно. Худалдан авалтын эцсийн шийдвэрийг борлуулалтын багтай баталгаажуулна.

## Аюулгүй байдал ба зардал

- Cloudflare Workers AI binding хэрэглэж буй тул token/API key GitHub frontend дээр тавихгүй.
- Browser Origin зөвшөөрөл зөвхөн `https://choijamtsdo-cell.github.io` байна. Domain өөрчлөгдвөл backend `ORIGIN`-ийг бас шинэчилнэ.
- Origin шалгалт нь скрипт, curl зэрэг non-browser хүсэлтийг бүрэн хаахгүй; олон хэрэглэгчтэй болохоос өмнө Cloudflare **Rate Limiting**, Bot protection/Turnstile болон request monitoring тохируулахыг зөвлөж байна.
- Код асуултын уртыг 500 тэмдэгт, харилцан яриаг сүүлийн 8 мессеж, хариултыг богино байдлаар хязгаарласан. Энэ нь зардлыг бүрэн хязгаарлах баталгаа биш.
- AI чат нь нэр, утас, хүсэлт Google Sheets-д **шууд бичихгүй**; хувийн мэдээллийг чатад оруулахгүй байхыг хэрэглэгчид сануулна. Жинхэнэ хүсэлтийн бүртгэл сайтад байсан `#quote` формоор ажиллана.
- Cloudflare-ийн 2026 оны үнийн мэдээллээр Workers AI Free allocation нь **10,000 Neurons/өдөр**, хэтэрсэн тохиолдолд Free төлөвлөгөөнд inference амжилтгүй болно; Paid нөхцөл өөр. Тариф өөрчлөгдөж болох тул Cloudflare Dashboard шалгана.

## Туршилт

1. Worker URL-ийн `/health` үр дүн `ok:true` байх.
2. Browser дээр сайтаа нээж AI товч харагдах.
3. «1 м³ хэдэн төгрөг вэ?» гэж асуухад **300,000₮/м³ жишиг үнэ** гэдгийг хэлж байгаа эсэх.
4. «Хүргэлттэй юу?» гэж асуухад хүргэнэ гэж **амлахгүй** байх.
5. «200 мм блок 1 м³-д хэдэн ширхэг вэ?» асуултад ойролцоогоор **34.7 ширхэг** гэж хариулах.
6. «Үнийн санал авах» товч `#quote` хэсэгт очих.
7. Mobile хэмжээ, хаах/нээх, Enter/Shift+Enter, алдааны fallback шалгах.
