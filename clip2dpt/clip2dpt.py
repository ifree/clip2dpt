#!/usr/bin/env python3

# fet clipboard image
from PIL import Image, ImageGrab, ImageOps
from io import BytesIO
import tempfile, argparse, os, sys, time
from dptrp1.dptrp1 import DigitalPaper, find_auth_files


DEST_FILE_NAME = "Document/clipboard.pdf"

def clipboard2pdf(resolution = (1650, 2200), inverse_color = False, flip = False) -> str:
    """Get clipboard image and save to pdf file.

    Returns:
        str: pdf file name
    """
    im = ImageGrab.grabclipboard()    
    
    if type(im) == list:
        im = Image.open(im[0])
    if im is None:
        return None
    
    # flip image to landscape mode, transpose + expand
    # if im.size[0] < im.size[1]: # portrait mode        
    #     im = im.rotate(90, expand=False)
    #     #im = ImageOps.exif_transpose(im)
        
    if flip:
        im = im.rotate(270, expand=False)

    # inverse color
    if inverse_color:
        im = ImageOps.invert(im)

    im.thumbnail(resolution, Image.Resampling.LANCZOS)

    # save pdf to temp file
    pdf_file = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
    pdf_file.close()
    im.save(pdf_file.name, 'PDF', title='clip2dpt')
    im.close()
    return pdf_file.name

def file2dpt(dp, args):
    """Send pdf file to Digital Paper."""

    pdf_file = clipboard2pdf(inverse_color=args.inv_color, flip=args.flip)
    if pdf_file is None or not os.path.exists(pdf_file):
        print("Could not get clipboard image.")
        exit(1)

    if not os.path.exists(pdf_file):
        print("Could not find pdf file.")
        return False
    dp.upload_file(pdf_file, DEST_FILE_NAME)
    # sleep 1s
    time.sleep(1)
    info = dp.list_document_info(DEST_FILE_NAME)
    dp.display_document(info["entry_id"])
    
    if args.keep_temp_file:
        os.startfile(pdf_file)
    else:
        os.remove(pdf_file)
        
    return True

def dpt2clip(dp):
    """Get image from Digital Paper and save to clipboard."""
    jpg_bytes = dp.take_screenshot()
    if jpg_bytes is None:
        return None
    im = Image.open(BytesIO(jpg_bytes))
    if im is None:
        return None
    if sys.platform.startswith("win"):
        with BytesIO() as output:
            im.save(output, "BMP")
            data = output.getvalue()[14:]
            import win32clipboard
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            win32clipboard.CloseClipboard()
            print('clipboard image saved.')
    else:
        print("Not supported platform.")        

def main():
    args = argparse.ArgumentParser(description='Send clipboard image to Digital Paper.')
    args.add_argument('--inv_color', '-i', help='inverse color', action='store_true', default=False)
    args.add_argument('--keep_temp_file', '-k', help='keep pdf file', action='store_true', default=False)
    args.add_argument('--flip', '-f', help='flip image', action='store_true', default=False)
    args.add_argument('--grab', '-g', help='grab image from Digital Paper', action='store_true', default=False)
    args = args.parse_args()

    dp = DigitalPaper()

    found_deviceid, found_privatekey = find_auth_files()
    if not os.path.exists(found_deviceid) or not os.path.exists(found_deviceid):
        print("Could not read device identifier and private key.")
        print("Please register with `dpt-rp1-py` frist.")
        exit(1)
    with open(found_deviceid) as fh:
        client_id = fh.readline().strip()
    with open(found_privatekey, "rb") as fh:
        key = fh.read()
    dp.authenticate(client_id, key)

    try:
        if args.grab:
            dpt2clip(dp)
        else:
            file2dpt(dp, args)

    except Exception as e:
        print("An error occured:", e, file=sys.stderr)
        print("For help, call:", sys.argv[0], "help")
        sys.exit(1)

if __name__ == '__main__':
    main()