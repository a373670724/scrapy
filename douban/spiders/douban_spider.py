# -*- coding: utf-8 -*-
import scrapy
from douban.items import DoubanItem
from openpyxl import Workbook

class DoubanSpiderSpider(scrapy.Spider):
    i=0
    rows = []
    #爬虫名
    name = 'douban_spider'
    #允许的域名
    allowed_domains = ['movie.douban.com']
    #入口URL，扔到调度器里
    start_urls = ['https://movie.douban.com/top250']
    book = Workbook()
    sheet = book.active
    # i=0

    def parse(self, response):
        # global i
        movie_list=response.xpath("//div[@class='article']//ol[@class='grid_view']/li")

        for i_item in movie_list:
            # print(i_item)
            douban_item=DoubanItem()
            douban_item['serial_number']=i_item.xpath(".//div[@class='item']//em/text()").extract_first()
            douban_item['movie_name'] = i_item.xpath(".//div[@class='info']/div[@class='hd']/a/span[1]/text()").extract_first()
            content = i_item.xpath(
                ".//div[@class='info']/div[@class='bd']/p[1]/text()").extract()
            content_s=''
            for i_content in content:
                content_s+="".join(i_content.split())
                douban_item['introduce']=content_s
            douban_item['star']=i_item.xpath( ".//span[@class='rating_num']/text()").extract_first()
            douban_item['evaluate']=i_item.xpath(".//div[@class='star']//span[4]/text()").extract_first()
            douban_item['describe']=i_item.xpath(".//p[@class='quote']/span[1]/text()").extract_first()
            DoubanSpiderSpider.rows.insert(DoubanSpiderSpider.i,(douban_item['serial_number'],douban_item['movie_name'],douban_item['introduce'],
                         douban_item['star'], douban_item['evaluate'], douban_item['describe'])
                         )
            DoubanSpiderSpider.i=DoubanSpiderSpider.i+1
            yield douban_item
        next_link=response.xpath(".//span[@class='next']/link/@href").extract()
        if next_link:
            next_link=next_link[0]
            yield scrapy.Request("https://movie.douban.com/top250"+next_link,callback=self.parse)
        if len(DoubanSpiderSpider.rows)==250:
            for row in DoubanSpiderSpider.rows:
                DoubanSpiderSpider.sheet.append(row)

                DoubanSpiderSpider.book.save('d:\\appending.xlsx')
            # print(douban_item)
