"""
Extract all images from a website xiaohongshu, and then combine into a docx file.
"""

import argparse
from src.utils import create_doc, create_image_dict, add_image_from_web, save_doc

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", "-u", default=None,
                        help="xiaohongshu website")
    parser.add_argument("--output", "-o", default="毛衣",
                        help="Output document name, only prefix is enough")
    args = parser.parse_args()
    
    if args.url is not None:
        images = create_image_dict(args.url)
        doc = create_doc()
        for k,v in images.items():
            add_image_from_web(doc,v)
        save_doc(doc,args.output)