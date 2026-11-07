import os
import json
import asyncio

from aiogram import Bot, Dispatcher, types
from aiogram.filters import CommandStart
from aiogram.types import BufferedInputFile
from openai import AsyncOpenAI
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

BOT_TOKEN = os.getenv("BOT_TOKEN")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
ai = AsyncOpenAI(api_key=OPENAI_API_KEY)


@dp.message(CommandStart())
async def start(message: types.Message):
    await message.answer(
        "Assalomu alaykum! Slayd Maker botiga xush kelibsiz!\n\n"
        "Taqdimot tayyorlash uchun mavzuni yuboring."
    )


@dp.message()
async def make_slides(message: types.Message):
    topic = message.text

    if not topic:
        await message.answer("Iltimos, mavzuni matn shaklida yuboring.")
        return

    status = await message.answer("⏳ Taqdimotingiz tayyorlanmoqda...")

    try:
        response = await ai.chat.completions.create(
            model="gpt-4o-mini",
            response_format={"type": "json_object"},
            messages=[
                {
                    "role": "system",
                    "content": (
                        "Professional taqdimot tuzuvchi mutaxassissan. "
                        "O'zbek tilida aynan 6 ta slayd tayyorla. "
                        "JSON formatida title va slides qaytar. "
                        "Har bir slaydda heading va content bo'lsin. "
                        "content qisqa, tushunarli va mazmunli bo'lsin."
                    )
                },
                {
                    "role": "user",
                    "content": f"Mavzu: {topic}"
                }
            ]
        )

        data = json.loads(response.choices[0].message.content)

        prs = Presentation()
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)

        for index, item in enumerate(data["slides"]):
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            background = slide.background.fill
            background.solid()
            background.fore_color.rgb = RGBColor(20, 35, 65)

            title_box = slide.shapes.add_textbox(
                Inches(0.7), Inches(0.5), Inches(12), Inches(1.2)
            )
            title = title_box.text_frame
            title.text = item["heading"]
            title.paragraphs[0].font.size = Pt(30)
            title.paragraphs[0].font.bold = True
            title.paragraphs[0].font.color.rgb = RGBColor(255, 210, 80)

            body_box = slide.shapes.add_textbox(
                Inches(0.9), Inches(2), Inches(11.5), Inches(4.5)
            )
            body = body_box.text_frame
            body.word_wrap = True
            body.text = item["content"]

            for paragraph in body.paragraphs:
                paragraph.font.size = Pt(23)
                paragraph.font.color.rgb = RGBColor(255, 255, 255)

        output = __import__("io").BytesIO()
        prs.save(output)

        file = BufferedInputFile(
            output.getvalue(),
            filename="Slayd_Maker.pptx"
        )

        await message.answer_document(
            file,
            caption="✅ Taqdimotingiz tayyor!"
        )

        await status.delete()

    except Exception as e:
        await status.edit_text(
            "Xatolik yuz berdi. Keyinroq qayta urinib ko'ring."
        )
        print("Xatolik:", e)


async def main():
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
