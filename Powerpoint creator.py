import requests
import re
from bs4 import BeautifulSoup
import os
from urllib.parse import urljoin
from pptx import Presentation
from pptx.util import Inches, Pt
from PIL import Image
import wikipedia
from openai import OpenAI

#Get rid of commas and square brackets in bullet points (change prompts)
#2 photos sometimes on page
#Make formatting better on slides (fix bullet points so slides filled out)
#Add image captions
#Add speaker notes (maybe use second ai agent)

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

def searchWikipedia():
    topicChosen = False

    while not topicChosen:
        topic = input("Enter a topic for your PowerPoint: ")

        # 2. Get the search results
        try:
            # Use search() to get a list of the top 5 matching titles
            results = wikipedia.search(topic, results=5)
            
            if not results:
                print(f"\nTopic not found!")
                return

            print(f"\n--- Select from the following! ---")
            
            # 3. Print the titles
            for i, title in enumerate(results):
                print(f"[{i+1}] {title}")
                #print(f"https://en.wikipedia.org/wiki/{title.replace(' ','_')}")
            
            print("-------------------------------------------------")
            choice = ''
            while choice not in {1,2,3,4,5}:
                choice = int(input('Enter topic number - '))
            topic = results[choice-1]

            topicChosen = True

        except wikipedia.exceptions.DisambiguationError:
            # This error is less likely when just searching, but good practice to keep
            print(f"\nThe topic '{topic}' is too general. Please be more specific.")
        except Exception as e:
            # A basic catch-all for other network or library errors
            print(f"\nAn unexpected error occurred: {e}")

    # make parent folder
    parentFolder = topic.replace(' ','_')
    os.makedirs(parentFolder, exist_ok=True)

    url = "https://en.wikipedia.org/wiki/"+parentFolder

    return url, topic, parentFolder

def removeTags(html):
    soup = BeautifulSoup(html, 'html.parser')
    visibleText = []

    for element in soup.find_all('p'):
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

url,topic,parentFolder = searchWikipedia()

print(url)
headers = {"User-Agent": "Mozilla/5.0"}
response = requests.get(url, headers=headers)
html = response.text

subTopicTitles = []
subTopicBodies = []
subTopicImages = []

avoidedContents = ['Citations','Notes','See also','References','References 2','Sources','Further reading','External links','Gallery']
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

#print(html)

contents = html.split('<div class="mw-heading mw-heading2"><h2 id="')
#splitting each html headline at id to get subtopics
contents.pop(0)

for content in contents:
    subTopicTitle = str((content.split('"'))[0]).replace('_',' ')
    #splitting end of id tag to get subtopic title

    if subTopicTitle not in avoidedContents:
        soup = BeautifulSoup(content, 'html.parser')
        imgs = []         # saved local file paths
        captions = []     # captions for saved images
        imagesSaved = 0

        # prepare folder for this subtopic
        topicFolder = os.path.join(parentFolder, subTopicTitle.replace(' ', '_'))
        os.makedirs(topicFolder, exist_ok=True)

        headers = {'User-Agent': 'Mozilla/5.0'}

        for img in soup.find_all('img'):
            if imagesSaved >= 2:
                break

            # pick candidate URL from src/srcset/data-src etc.
            src = img.get('src') or img.get('data-src') or img.get('data-image-src') or img.get('srcset') or ''
            if not src:
                continue

            # handle srcset lists like "url1 1x, url2 2x" - pick the last candidate url
            if ',' in src and ' ' in src:
                candidates = [s.strip() for s in src.split(',') if s.strip()]
                last = candidates[-1]
                src = last.split()[0]

            # normalise to absolute URL
            if src.startswith('//'):
                link = 'https:' + src
            elif src.startswith('/'):
                link = urljoin('https://en.wikipedia.org', src)
            elif src.startswith('http'):
                link = src
            else:
                continue

            # basic filename and filter check
            url_path = link.split('?', 1)[0]
            filename_only = os.path.basename(url_path)
            if any(bad in filename_only for bad in avoidedImages):
                continue

            # caption heuristics
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

            ext = '.png'

            local_name = f"img{imagesSaved}{ext}"
            local_path = os.path.join(topicFolder, local_name)

            # download and save
            try:
                r = requests.get(link, headers=headers, stream=True, timeout=12)
                r.raise_for_status()
                with open(local_path, 'wb') as f:
                    for chunk in r.iter_content(8192):
                        if chunk:
                            f.write(chunk)
                print(f"Saved: {local_path}")
                imgs.append(local_path)
                captions.append(cap)
                imagesSaved += 1
            except Exception as e:
                print(f"Failed to save {link}: {e}")
                # try next image tag

        # at this point imgs and captions have up to 2 saved items each
        if not imgs:
            print(f"No images saved for subtopic: {subTopicTitle}")

        subTopicImages.append(list(zip(imgs, captions)))

        subTopicContent = re.sub(r'\[[^\]]*\]', '', removeTags(content)) #get just p tags and get rid of square brackets
        subTopicTitles.append(subTopicTitle)
        subTopicBodies.append(subTopicContent)

        print(subTopicTitle)
        print(subTopicContent[:1000])
        print(subTopicImages)

#Make powerpoint:

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
        fontSize = 14 #smaller font if image on slide
        if leftText:
            subtitle = slide.placeholders[1]
            slide.shapes._spTree.remove(slide.placeholders[2]._element)
        else:
            subtitle = slide.placeholders[2]
            slide.shapes._spTree.remove(slide.placeholders[1]._element)
    else:
        fontSize = 24 #larger text if no images
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

        localImagePath = os.path.join(topicFolder, 'img0.png')   # simple guess
        width, height, widthLost, heightLost = getImageSize(localImagePath)

        image = localImagePath
        width, height, widthLost, heightLost = getImageSize(image)

        if leftText:
            widthFromLeft = 5
        else:
            widthFromLeft = 0.5
        
        slide.shapes.add_picture(image, Inches(widthFromLeft+widthLost/2), Inches(1.5+heightLost/2), Inches(width), Inches(height))
            
        '''
        caption_top = 1.5+heightLost/2 + (pic.height.inches) + 0.1  # add a small gap

        # Add text box for caption
        textbox = slide.shapes.add_textbox(Inches(left), Inches(caption_top), Inches(width), Inches(0.5))
        text_frame = textbox.text_frame
        text_frame.text = caption_text or ""
        text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        text_frame.paragraphs[0].font.size = Pt(12)
        text_frame.paragraphs[0].font.italic = True
        '''
    #caption

    leftText = not leftText

# before presentation.save(powerpointName)
os.makedirs(os.path.dirname(powerpointName), exist_ok=True)
presentation.save(powerpointName)
os.startfile(powerpointName)

'''
#Run powerpoint
presentation.save(powerpointName)
os.startfile(powerpointName)
	'''

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
