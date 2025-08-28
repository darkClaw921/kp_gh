import requests
from fast_bitrix24 import BitrixAsync
from pprint import pprint
from dotenv import load_dotenv
import os
from dataclasses import dataclass
import asyncio


class Product:
    id: int
    name: str = 'name'
    price: str='PRICE'
    quantity: int


class Deal:
    id: int
    opportunity: str ='OPPORTUNITY'
    obem_po_porametram: str = 'UF_CRM_1756231768065'
    frakcia: str = 'UF_CRM_1756231794903'
    ypakovka: str = 'UF_CRM_1756232055022'
    dostavka: str = 'UF_CRM_1756232017298'
    fileKP: str = 'UF_CRM_1756392242954'

# [{"NAME":"не выбрано","VALUE":"","IS_SELECTED":false},{"NAME":"сетка","VALUE":556,"IS_SELECTED":false},{"NAME":"мешки","VALUE":558,"IS_SELECTED":true},{"NAME":"биг-бэги","VALUE":560,"IS_SELECTED":false}]
# [{"NAME":"не выбрано","VALUE":"","IS_SELECTED":false},{"NAME":"без учета доставки","VALUE":552,"IS_SELECTED":false},{"NAME":"с учетом доставки","VALUE":554,"IS_SELECTED":true}]

class TypeYpakovka:
    no_selected:str= ''
    setka:str= '556'
    meshki:str= '558'
    big_bagi:str= '560'


class TypeDostavka:
    no_selected:str= ''
    bez_dostavki:str= '552'
    s_dostavki:str= '554'

typeYpakovka={
    '':'Не выбрано',
    '556':'сетка',
    '558':'мешки',
    '560':'биг-бэги',
}

typeDostavka={
    '':'Не выбрано',
    '552':'без учета доставки',
    '554':'с учетом доставки',
}



load_dotenv()
WEBHOOK=os.getenv('WEBHOOK')

bit = BitrixAsync(WEBHOOK)


async def get_product(product_id) -> dict:
    product = await bit.call('catalog.product.get', {'id': product_id})
    return product

async def get_deal(deal_id) -> dict:
    deal = await bit.call('crm.deal.get', {'id': deal_id})
    if isinstance(deal,dict):
        if deal.get('order0000000000') is not None:
            deal=deal['order0000000000']
    return deal

async def get_deal_products(deal_id) -> list:
    products = await bit.call('crm.deal.productrows.get', {'id': deal_id})
    return products

async def download_images(url:str, namefield:str,productId:int,fileId:int):
    """
    Делает POST как в curl к catalog.product.download и сохраняет файл.
    Возвращает путь к сохранённому файлу или None.
    """
    import asyncio

    # main_dir = 'images'
    # os.makedirs(main_dir, exist_ok=True)

    endpoint = f"{WEBHOOK}/catalog.product.download"
    headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
    }
    payload = {
        'fields': {
            'fileId': fileId,
            'productId': productId,
            'fieldName': namefield,
        }
    }

    def _post_and_save():
        resp = requests.post(endpoint, headers=headers, json=payload)
        if resp.status_code != 200:
            print(f"Ошибка загрузки файла: {resp.status_code} {resp.text}")
            return None
        content_type = resp.headers.get('Content-Type', '')
        if 'image/png' in content_type or '.png' in url:
            ext = '.png'
        elif 'image/jpeg' in content_type or 'image/jpg' in content_type:
            ext = '.jpg'
        else:
            ext = '.png'
        # file_name = f"{main_dir}/{productId}_{namefield}{ext}"
        file_name = f"{productId}_{namefield}{ext}"
        with open(file_name, 'wb') as f:
            f.write(resp.content)
        return file_name

    return await asyncio.to_thread(_post_and_save)
    
async def upload_file_to_deal(deal_id,file_name):
    import base64

    with open(file_name, "rb") as file:
        encoded_file = base64.b64encode(file.read()).decode('utf-8')

    # deal= await get_deal(deal_id)
    
    fields={
        Deal.fileKP: {"fileData": [os.path.basename(file_name), str(encoded_file)]}
    }
    await bit.call('crm.deal.update', {'ID': deal_id, 'FIELDS': fields})


async def get_all_info(deal_id):
    # deal_id=8076
    deal= await get_deal(deal_id)
    print(deal)
    frakcia=deal[Deal.frakcia]
    ypakovka=typeYpakovka[deal[Deal.ypakovka]]
    obem_po_porametram=deal[Deal.obem_po_porametram]

    if typeDostavka[deal[Deal.dostavka]] in ['с учетом доставки']:

        dostavka=True
    else:
        dostavka=False

    opportunity=deal[Deal.opportunity]

    print(frakcia,ypakovka,dostavka,opportunity)
    products= await get_deal_products(deal_id)
    pprint(products)
    product= await get_product(products['PRODUCT_ID'])
    pprint(product)

    productName=product[Product.name]
    productPrice=products[Product.price]
    # print(f'productName: {productName}')
    # print(f'productPrice: {productPrice}')
    images={'default': 'path', 'dry': 'path', 'wet': 'path', 'lit': 'path'}
    imagesPreviewPicture = await download_images(url=product['previewPicture']['url'],
                                   namefield='previewPicture',
                                   productId=product['id'],
                                   fileId=product['previewPicture']['id'])
    
    images['default']=imagesPreviewPicture
    images['dry']=imagesPreviewPicture
    images['wet']=imagesPreviewPicture
    images['lit']=imagesPreviewPicture
    # print(images)

    print(f'frakcia: {frakcia}')
    print(f'ypakovka: {ypakovka}')
    print(f'dostavka: {dostavka}')
    print(f'opportunity: {opportunity}')
    print(f'productName: {productName}')
    print(f'productPrice: {productPrice}')
    print(f'images: {images}')
    # images
    return frakcia,ypakovka,dostavka,opportunity,productName,images,productPrice,obem_po_porametram

if __name__ == '__main__':
    asyncio.run(get_all_info(8076))





