import requests
from bs4 import BeautifulSoup
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from PIL import Image
from openai import OpenAI

ai = OpenAI(api_key=open("openAI_key.txt").read().strip())

def makeBulletPoints(visibleText):
    # Prepare the assistant call to produce strict JSON output:
    completion = ai.chat.completions.create(
        model='gpt-3.5-turbo',
        temperature=0.0,            # deterministic
        max_tokens=400,
        messages=[
            {
                'role': 'system',
                'content': (
                    "You are an expert slide-writer that converts a chunk of text into "
                    "concise PowerPoint bullet points. STRICT RULES - follow them exactly:\n"
                    "1) Output only valid JSON (no surrounding text). The JSON must be a "
                    "single array of strings, e.g. [\"Bullet 1\",\"Bullet 2\"].\n"
                    "2) Produce between 3 and 8 bullets. Never exceed 10 bullets.\n"
                    "3) TOTAL output length must not exceed 150 words and 800 characters.\n"
                    "4) Each bullet should be a single short sentence or phrase, ideally 6-15 words.\n"
                    "5) Do not include citations, bracketed references, html, or source text.\n"
                    "6) Use plain text only; do not return markdown, lists, headings or extra fields.\n"
                    "7) If the input is short or has too little content, still return 3 concise bullets.\n"
                    "8) If you cannot identify 3 meaningful bullets, return the three best short summary phrases.\n"
                    "Tone: neutral, factual, slide-friendly.\n"
                    "Example input -> output:\n"
                    "Input: 'The Hindenburg disaster occurred in 1937 when the German passenger airship LZ 129 Hindenburg caught fire while docking in New Jersey.'\n"
                    "Output: [\"Hindenburg disaster: LZ 129 caught fire while docking (1937)\", "
                    "\"Major loss of life and media coverage\", \"Fire highlighted hydrogen safety risks\"]"
                )
            },
            {
                'role': 'user',
                'content': (
                    "Convert the following text into slide-ready bullets as a JSON array of strings. "
                    "Remember the strict limits above.\n\n"
                    f"Text: {visibleText}"
                )
            }
        ]
    )

    bulletPoints = completion.choices[0].message.content
    return bulletPoints

topic = input('Enter any topic - ')
topic = topic.capitalize()

# make parent folder
parentFolder = topic.replace(' ','_')
os.makedirs(parentFolder, exist_ok=True)

def removeTags(html):
    soup = BeautifulSoup(html, 'html.parser')
    visibleText = []

    for element in soup.find_all(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'span']):
        visibleText.append(element.get_text())

    return ' '.join(visibleText)

def getImageSize(image):
    imWidth, imHeight = Image.open(image).size
    dimensions = imWidth/imHeight
    maxWidth = 4.5
    maxHeight = 5.5
    maxDimension = maxWidth/maxHeight
    if dimensions>maxDimension: #landscape
        width = maxWidth
        height = maxWidth*(1/dimensions)
    else: #portrait
        height = maxHeight
        width = maxHeight*dimensions
    heightLost = maxHeight-height
    widthLost = maxWidth-width

    return width, height, widthLost, heightLost


url = "https://en.wikipedia.org/wiki/"+parentFolder
headers = {"User-Agent": "Mozilla/5.0"}
response = requests.get(url, headers=headers)
html = response.text

# split at each h2 section
subTopicHTML = html.split('<div class="mw-heading mw-heading2"><h2 id="')[1:]

subTopicTitles = []
subTopicBodies = []
subTopicImages = []

avoidedContents = ['Citations','Notes','See also','References','Sources','Further reading','External links','Gallery']
avoidedImages = [
    'Question_book-new.svg',
    'Nuvola_apps_kaboodle.svg',
    'Ambox_current_red_Asia_Australia.svg',
    'Information_icon4.svg',
    'Climate_change_icon',
    'Symbol_list_class.svg',
    'Ambox_rewrite.svg',
    'Ambox_important.svg'
]

for part in subTopicHTML:
    # Split into title and body
    titleStart = part.find(">") + 1
    titleEnd = part.find("</h2>")
    title = part[titleStart:titleEnd]
    if "</span>" in title:
        title = title.split("</span>", 1)[1]
    bodyHTML = part[titleEnd + len("</h2>"):]

    if title not in avoidedContents:

        soup = BeautifulSoup(bodyHTML, 'html.parser')
        imgs = []
        captions = []

        for img in soup.find_all('img'):
            src = img.get('src') or img.get('data-src') or img.get('data-srcset') or img.get('srcset')
            if not src:
                continue

            if ',' in src and (' ' in src):
                candidates = [s.strip() for s in src.split(',') if s.strip()]
                last = candidates[-1]
                src = last.split()[0]

            if src.startswith('//'):
                link = 'https:' + src
            elif src.startswith('/'):
                link = 'https://en.wikipedia.org' + src
            elif src.startswith('http'):
                link = src
            else:
                continue

            filename = link.split('/')[-1].split('?')[0]
            print(filename)
            skipImage = False
            for bad in avoidedImages:
                if bad in filename:
                    skipImage = True
                    break
            if skipImage:
                continue


            cap = ""
            parentFig = img.find_parent('figure')
            if parentFig:
                capTag = parentFig.find('figcaption') or parentFig.find(class_='thumbcaption')
                if capTag:
                    cap = capTag.get_text(separator=' ', strip=True)
            if not cap:
                capTag = img.find_next('figcaption') or img.find_next(class_='thumbcaption')
                if capTag:
                    cap = capTag.get_text(separator=' ', strip=True)
            if not cap:
                cap = img.get('alt') or img.get('title') or ""

            imgs.append(link)
            captions.append(cap)

        body = removeTags(bodyHTML)

        subTopicTitles.append(title)
        subTopicBodies.append(body)
        subTopicImages.append(list(zip(imgs, captions)))

        if imgs:
            topicFolder = os.path.join(parentFolder, title.replace(' ', '_'))
            os.makedirs(topicFolder, exist_ok=True)

            for imgNo, link in enumerate(imgs):
                if link.startswith('//'):
                    link = 'https:' + link
                elif link.startswith('/'):
                    link = 'https://en.wikipedia.org' + link

                urlPath = link.split('?', 1)[0]
                ext = os.path.splitext(urlPath)[1].lower()
                if ext not in ('.jpg', '.jpeg', '.png', '.gif', '.svg'):
                    ext = '.jpg'

                filename = os.path.join(topicFolder, f"img{imgNo}{ext}")
                try:
                    r = requests.get(link, headers=headers, stream=True, timeout=15)
                    r.raise_for_status()
                    with open(filename, 'wb') as f:
                        for chunk in r.iter_content(8192):
                            if chunk:
                                f.write(chunk)
                    print(f"Saved: {filename}")
                except Exception as e:
                    print(f"Failed to save {link}: {e}")

#Create powerpoint
powerpointName = '{}\{} Powerpoint.pptx'.format(topic.replace(' ','_'),topic)
presentation = Presentation()
slide = presentation.slides.add_slide(presentation.slide_layouts[0]) #title layout
title = slide.shapes.title
subtitle = slide.placeholders[1]
    
title.text = str(topic)
if len(subTopicTitles) > 3:
    subtitleText = '{}, {}, {} and more'.format(*subTopicTitles[:3])
else:
    subtitleText = ', '.join(subTopicTitles)
#adding first few subtopics as title page subtitle

subtitle.text = subtitleText

leftText = True
for topicTitle, body, images in zip(subTopicTitles, subTopicBodies, subTopicImages):
    print("TITLE:", topicTitle)
    print("BODY:", body[:2000])
    print("IMAGES:")

    pictureSlide = bool(images)

    if pictureSlide:
        bullet_slide_layout = presentation.slide_layouts[3] #picture and text layout
    else:
        bullet_slide_layout = presentation.slide_layouts[1] #just text layout

    slide = presentation.slides.add_slide(bullet_slide_layout)

    if pictureSlide:
        fontSize = 12 #smaller font if image on slide
        if leftText:
            subtitle = slide.placeholders[1]
            slide.shapes._spTree.remove(slide.placeholders[2]._element)
        else:
            subtitle = slide.placeholders[2]
            slide.shapes._spTree.remove(slide.placeholders[1]._element)
    else:
        fontSize = 18 #larger text if no images
        subtitle = slide.placeholders[1]

    title = slide.shapes.title
    title.text = topicTitle
    visibleText = body
    slideContent = makeBulletPoints(visibleText) #uses ai to turn text into bullet points
    '''
    lines = slideContent.split('\n')
    linesWithoutDashes = [line.lstrip('- ') for line in lines]
    slideContent = '\n'.join(linesWithoutDashes)

    if len(slideContent)>1200:
        fontSize-=2
    fontSize = Pt(fontSize)
    '''
    slideContent = slideContent.replace('[','').replace(']','').replace('"','')
    subtitle.text = slideContent
    text_frame = subtitle.text_frame
    for paragraph in text_frame.paragraphs:
        font = paragraph.font
        font.size = Pt(fontSize)

    if pictureSlide:
        topicFolder = os.path.join(parentFolder, topicTitle.replace(' ', '_'))
        # assume you saved files as img0, img1 ... with original extension .jpg (or detect)
        try:
            localImagePath = os.path.join(topicFolder, 'img0.jpg')   # simple guess
            width, height, widthLost, heightLost = getImageSize(localImagePath)
        except:
            localImagePath = os.path.join(topicFolder, 'img0.png')   # simple guess
            width, height, widthLost, heightLost = getImageSize(localImagePath)

        image = localImagePath
        width, height, widthLost, heightLost = getImageSize(image)
        
        if leftText:
            widthFromLeft = 5
        else:
            widthFromLeft = 0.5
        slide.shapes.add_picture(image, Inches(widthFromLeft+widthLost/2), Inches(1.5+heightLost/2), Inches(width), Inches(height))
#caption

    leftText = not leftText

# before presentation.save(powerpointName)
os.makedirs(os.path.dirname(powerpointName), exist_ok=True)
presentation.save(powerpointName)
os.startfile(powerpointName)


#Run powerpoint
presentation.save(powerpointName)
os.startfile(powerpointName)
	

'''
# print summary
for title, body, images in zip(subTopicTitles, subTopicBodies, subTopicImages):
    print("TITLE:", title)
    print("BODY:", body[:2000])
    print("IMAGES:")

    if not images:
        print("  (none)")
    else:
        for link, caption in images:
            print("  -", link)
            if caption:
                print("    Caption:", caption)
    print()
'''
