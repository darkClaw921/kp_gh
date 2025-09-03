from json import load
import asyncio
from workBitrix import create_product

async def main(fileName='items.json'):
    with open(fileName, 'r') as f:
        items = load(f)

    for item in items:
        product_name=item['name']
        product_price=0
        product_quantity=1
        if item['image']==None:
            continue
        product_image=item['image_base64']
        
        await create_product(product_name, product_price, product_quantity, product_image, isBase64=False)
        # return 0

if __name__ == '__main__':
    filename='items2.json'
    asyncio.run(main(fileName=filename))

