// VIA ELITE — AI 모델컷 생성 백엔드 (Vercel 서버리스 함수)
// 브라우저에서 받은 상품 이미지 + 옵션으로 프롬프트를 조립하고 Gemini(나노바나나)를 호출한다.
// API 키는 Vercel 환경변수 GEMINI_API_KEY 에만 보관한다. (코드/프론트에 절대 노출 금지)

const MODEL = 'gemini-2.5-flash-image';
const ENDPOINT = 'https://generativelanguage.googleapis.com/v1beta/models/' + MODEL + ':generateContent';

// ----- 카테고리별 착용 방식 / 기본 구도 / 클로즈업 부위 -----
const CATEGORY = {
  clothing:  { wear: "actually WEARING this garment on their body (the garment is worn on the body, NOT held in the hands)", framing: "full-body or upper-body composition that clearly shows the garment worn on the body", closeup: "the part of the body where the garment is worn, the garment filling most of the frame" },
  bag:       { wear: "holding or carrying this bag naturally (in hand or on the shoulder)", framing: "upper-body composition with the bag clearly visible and in sharp focus", closeup: "the bag held in hand or on the shoulder, the bag filling most of the frame" },
  shoes:     { wear: "actually WEARING these shoes on their feet (the shoes are worn on the feet, NOT held)", framing: "full-body or lower-body composition that clearly shows the shoes worn on the feet", closeup: "the feet wearing the shoes, the shoes filling most of the frame, shot from a low side angle" },
  jewelry:   { wear: "actually WEARING this jewelry on the appropriate part of the body (e.g. neck, ear, wrist or finger), NOT holding it", framing: "close-up portrait composition focused on the body area where the jewelry is worn", closeup: "the body area wearing the jewelry (neck, ear, wrist or finger), the jewelry filling most of the frame" },
  accessory: { wear: "naturally wearing or using this accessory on the appropriate part of the body, NOT just holding it", framing: "upper-body composition with the accessory clearly visible and in sharp focus", closeup: "the body area where the accessory is worn, the accessory filling most of the frame" },
};

function shotFraming(shot, category, cat) {
  switch (shot) {
    case 'full':
      return "FULL-BODY composition showing the entire model head to toe, the product clearly visible";
    case 'half':
      if (category === 'shoes') return "LOWER-BODY composition (from the waist down) clearly showing the shoes worn on the feet, do NOT show the face";
      return "HALF-BODY composition tightly cropped to the body area where the product is actually worn — UPPER body (waist up) for tops, jackets and upper garments, LOWER body (waist down) for pants and lower garments — so the product is large and clearly visible; do NOT show the full body";
    case 'closeup':
      return "TIGHT CLOSE-UP shot focused on " + cat.closeup + " (do NOT show the full body)";
    default:
      return cat.framing; // auto
  }
}

function buildPrompt(opts) {
  const gender = opts.gender === 'man' ? 'man' : 'woman';
  const poss = opts.gender === 'man' ? 'his' : 'her';
  const ageMap = { '20s': poss + ' 20s', '30s': poss + ' 30s', '40s': poss + ' 40s', '50s+': poss + ' 50s or older' };
  const agePhrase = 'in ' + (ageMap[opts.age] || ageMap['30s']);
  const cat = CATEGORY[opts.category] || CATEGORY.clothing;
  const multi = opts.imageCount > 1
    ? "The provided images show the SAME single product from different angles — study them together to reproduce every detail accurately, including parts hidden in any one photo.\n\n"
    : "";
  return (
"First, look carefully at the provided product image(s) and identify exactly what the product is. " + multi +
"Generate a photorealistic editorial photograph of a Korean " + gender + " " + agePhrase + " " + cat.wear + ".\n\n" +
"CRITICAL — Keep the product IDENTICAL to the reference: same shape, color, material, texture, hardware, stitching, pattern and logo. Do not redesign, recolor, or alter the product in any way. Reproduce the exact same product, just shown on the model.\n\n" +
"Model: a sophisticated, attractive Korean " + gender + " " + agePhrase + " with a luxury-model look, elegant confident expression, refined editorial posture, flawless realistic skin and natural hands. Styling: high-fashion, expensive minimal styling in refined neutral tones that complements (does not compete with) the product.\n\n" +
"Scene & mood: ultra-premium LUXURY BRAND CAMPAIGN aesthetic, like a Vogue / high-fashion magazine editorial. Professional studio with sophisticated lighting — soft key light plus subtle rim light, elegant soft shadows. Refined upscale background (warm beige, soft gradient, or elegant neutral seamless), cinematic and aspirational mood, expensive and tasteful.\n\n" +
"Framing: " + shotFraming(opts.shot, opts.category, cat) + ", with the product as the clear focal point. Quality: ultra-high-end fashion photography, crisp fine detail, elegant cinematic color grading, sharp focus on the product, shallow depth of field, 85mm lens look, photorealistic, premium and aspirational.\n\n" +
"Avoid: cheap or amateur look, flat lighting, casual snapshot feel, extra logos or text, watermarks, distorted hands, busy background, holding the item if it should be worn, and any change to the product."
  );
}

async function callGemini(key, body, tries) {
  for (let i = 0; i < tries; i++) {
    let res, data;
    try {
      res = await fetch(ENDPOINT, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'x-goog-api-key': key },
        body: JSON.stringify(body)
      });
      data = await res.json();
    } catch (e) {
      if (i < tries - 1) { await new Promise(r => setTimeout(r, 1500 * (i + 1))); continue; }
      throw new Error('Gemini 연결 오류: ' + e.message);
    }
    if (res.ok) return data;
    const code = data && data.error && data.error.code;
    const status = data && data.error && data.error.status;
    const transient = code === 500 || code === 503 || status === 'INTERNAL' || status === 'UNAVAILABLE';
    if (transient && i < tries - 1) { await new Promise(r => setTimeout(r, 1800 * (i + 1))); continue; }
    throw new Error((data && data.error && data.error.message) ? data.error.message : ('HTTP ' + res.status));
  }
}

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'POST 요청만 허용됩니다.' });
    return;
  }
  const key = process.env.GEMINI_API_KEY;
  if (!key) {
    res.status(500).json({ error: '서버에 API 키(GEMINI_API_KEY)가 설정되지 않았습니다.' });
    return;
  }

  // 본문 파싱
  let body = req.body;
  if (typeof body === 'string') { try { body = JSON.parse(body); } catch (e) { body = {}; } }
  body = body || {};

  const images = Array.isArray(body.images) ? body.images : (body.image ? [body.image] : []);
  if (images.length === 0) {
    res.status(400).json({ error: '상품 이미지가 없습니다.' });
    return;
  }

  // 이미지 파트 구성 (data URL → inline_data)
  const parts = [];
  for (const dataUrl of images.slice(0, 2)) {
    const m = /^data:(.+?);base64,(.*)$/.exec(dataUrl || '');
    if (m) parts.push({ inline_data: { mime_type: m[1], data: m[2] } });
  }
  if (parts.length === 0) {
    res.status(400).json({ error: '이미지 형식이 올바르지 않습니다.' });
    return;
  }

  const prompt = buildPrompt({
    gender: body.gender, age: body.age, category: body.category, shot: body.shot,
    imageCount: parts.length
  });
  parts.push({ text: prompt });

  try {
    const data = await callGemini(key, { contents: [{ parts: parts }] }, 3);
    const respParts = (((data.candidates || [])[0] || {}).content || {}).parts || [];
    const imgPart = respParts.find(p => (p.inlineData && p.inlineData.data) || (p.inline_data && p.inline_data.data));
    if (!imgPart) {
      const txt = respParts.map(p => p.text).filter(Boolean).join(' ');
      res.status(502).json({ error: '이미지가 반환되지 않았습니다.' + (txt ? ' (' + txt.slice(0, 160) + ')' : '') });
      return;
    }
    const inl = imgPart.inlineData || imgPart.inline_data;
    const mime = inl.mimeType || inl.mime_type || 'image/png';
    res.status(200).json({ image: 'data:' + mime + ';base64,' + inl.data });
  } catch (err) {
    res.status(502).json({ error: err.message });
  }
};
