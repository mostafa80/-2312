import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import { GoogleGenAI, Type } from '@google/genai';
import { createServer as createViteServer } from 'vite';
import { parseLocalWrittenCommand } from './src/utils/excelEngine';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json({ limit: '10mb' }));

  app.post('/api/ai-command', async (req, res) => {
    const { commandText, companies = [], sampleItems = [] } = req.body ?? {};

    if (!commandText || typeof commandText !== 'string') {
      res.status(400).json({ error: 'يرجى كتابة الأمر المطلوب تنفيذه.' });
      return;
    }

    const apiKey = process.env.GEMINI_API_KEY;
    const hasValidKey =
      Boolean(apiKey) &&
      apiKey !== 'MY_GEMINI_API_KEY' &&
      apiKey!.trim().length > 10;

    if (hasValidKey) {
      try {
        const ai = new GoogleGenAI({
          apiKey: process.env.GEMINI_API_KEY,
          httpOptions: {
            headers: {
              'User-Agent': 'aistudio-build',
            },
          },
        });

        const systemInstruction = `أنت مساعد ذكي متخصص في تنظيم شيتات الإكسيل المجمعة لعدة شركات وبديل ذكي لمعادلات Excel.
المستخدم يرفع شيت إكسيل يحتوي على عدة شركات وأصناف وأسعار وأرصدة دفترية وفعلية وفروقات وحالات (عجز، زيادة، مطابق، تالف، مرتجع).
أعمدة الجدول في الإكسيل هي:
A: الشركة (company)
B: الكود (code)
C: اسم الصنف (itemName)
D: التصنيف (category)
E: سعر الوحدة (unitPrice)
F: الرصيد الدفتري (bookQty)
G: الجرد الفعلي (actualQty)
H: الفارق بالعدد (variance = G - F)
I: قيمة الفارق المالية (varianceValue = H * E)
J: الحالة (status: عجز | زيادة | مطابق | تالف | مرتجع)

الشركات الموجودة حالياً في الشيت: ${JSON.stringify(companies)}
عينة من الأصناف: ${JSON.stringify(sampleItems.slice(0, 5))}

حول طلب المستخدم المكتوب باللغة العربية إلى قائمة عمليات منظمة (operations) يتم تطبيقها فوراً على الجدول:
أنواع العمليات المتاحة (type):
- PURGE_ZERO_AND_MATCHED: حذف الأصناف المطابقة أو التي يكون الفارق فيها صفر (0).
- ORGANIZE_PRECISE: تنظيم وترتيب شامل ودقيق للشيت (دمج المكرر + حذف المطابق والفارق صفر + ترتيب الشركات حسب الحالة والسعر).
- SORT_ROWS: ترتيب الصفوف حسب field (company | itemName | unitPrice | bookQty | actualQty | variance | absVariance | varianceValue | status) و direction (asc | desc).
- FILTER_KEEP_COMPANY: إبقاء أصناف شركة معينة فقط في الجدول (حدد companyName).
- DELETE_COMPANY: حذف أصناف شركة معينة (حدد companyName).
- DELETE_BY_STATUS: حذف حالة معينة (حدد status: عجز | زيادة | مطابق | تالف | مرتجع).
- KEEP_ONLY_STATUS: عرض حالة معينة فقط (حدد status).
- DELETE_BY_PRICE: حذف الأصناف حسب شرط سعر (حدد operator: lt | gt | lte | gte | eq و value).
- KEEP_BY_PRICE: إبقاء الأصناف حسب شرط سعر (حدد operator و value).
- ADJUST_PRICE_PERCENT: تعديل الأسعار بنسبة مئوية موجبة للزيادة أو سالبة للخصم (حدد value واختيارياً companyName).
- DEDUPLICATE_ITEMS: دمج الأصناف المكررة لنفس الشركة وجمع كمياتها.
- RECALCULATE_STATUSES: إعادة احتساب الفروقات والحالات تلقائياً.
- RESTORE_PURGED: استرجاع الأصناف المحذوفة/المستبعدة.
- ADD_FORMULA_COLUMN: إضافة عمود معادلة إكسيل جديد للجدول. استخدم {row} كرقم الصف في formula مثل "=ROUND(E{row}*1.14, 2)" أو "=ROUND(G{row}*E{row}, 2)" وحدد columnTitle.
- EVALUATE_FORMULA: حساب معادلة إكسيل إجمالية مثل "=SUM(I:I)" أو "=AVERAGE(E:E)" أو "=SUMIF(J:J, \\"عجز\\", I:I)" أو "=COUNTIF(J:J, \\"عجز\\")" مع وضع وصف النتيجة في explanation.`;

        const response = await ai.models.generateContent({
          model: 'gemini-3.8-flash',
          contents: commandText,
          config: {
            systemInstruction,
            responseMimeType: 'application/json',
            responseSchema: {
              type: Type.OBJECT,
              properties: {
                operations: {
                  type: Type.ARRAY,
                  items: {
                    type: Type.OBJECT,
                    properties: {
                      type: { type: Type.STRING },
                      field: { type: Type.STRING },
                      direction: { type: Type.STRING },
                      companyName: { type: Type.STRING },
                      status: { type: Type.STRING },
                      operator: { type: Type.STRING },
                      value: { type: Type.NUMBER },
                      columnTitle: { type: Type.STRING },
                      formula: { type: Type.STRING },
                      explanation: { type: Type.STRING },
                    },
                    required: ['type'],
                  },
                },
                assistantMessage: {
                  type: Type.STRING,
                },
              },
              required: ['operations', 'assistantMessage'],
            },
          },
        });

        const rawJson = response.text?.trim();
        if (rawJson) {
          const parsed = JSON.parse(rawJson);
          if (Array.isArray(parsed.operations) && parsed.operations.length > 0) {
            res.json({
              operations: parsed.operations,
              assistantMessage: parsed.assistantMessage,
              source: 'gemini-ai',
            });
            return;
          }
        }
      } catch (err) {
        console.error('Gemini command interpretation fallback to local engine:', err);
      }
    }

    // Fast deterministic fallback if offline or key not set
    const localOps = parseLocalWrittenCommand(commandText, companies);
    if (localOps && localOps.length > 0) {
      res.json({
        operations: localOps,
        assistantMessage: 'تم فهم الأمر وتنفيذه فوراً عبر المحرك الذكي.',
        source: 'local-smart',
      });
      return;
    }

    // Default helpful organization if command was very general
    res.json({
      operations: [{ type: 'ORGANIZE_PRECISE' }],
      assistantMessage:
        'تم تنفيذ التنظيم الدقيق للشيت (حذف المطابق والفارق الصفري وترتيب الشركات حسب الحالة والسعر).',
      source: 'local-smart',
    });
  });

  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(__dirname, 'dist');
    app.use(express.static(distPath));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
