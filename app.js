const TelegramBot = require('node-telegram-bot-api');
const moment = require('moment');
const fs = require('fs');
const writeXlsxFile = require('write-excel-file/node')
const pg = require('pg');
const path = require('path');
const { Client } = pg
const token = '7387625761:AAGBsl5fo5-0_bld8Mq_IlJ1AgVEjoEcOHg';

const client = new Client({
    user: 'sweet-and-tasty-ro',
    password: '.VMXsBAqYLsva-m3T!P6',
    host: '195.49.212.216',
    port: 5432,
    database: 'sweet-and-tasty',
})

client.connect().then(() => {
    const bot = new TelegramBot(token, {polling: true});

    bot.on('message', (msg) => {
      const chatId = msg.chat.id;

      if (msg.text.toLowerCase().includes('катю') || msg.text.toLowerCase().includes('катя') || msg.text.toLowerCase().includes('кати')) {
        client.query(`
            select
            distinct p.name as product_name,
            (
                select sum(oi.quantity)
                from orders_orderitem oi
                join orders_order oo on oo.id = oi.order_id
                where oi.product_id = p.id and
                cast(oo.created_at as TIMESTAMP) > cast(concat(CURRENT_DATE, ' 00:00:00') as TIMESTAMP) and
                cast(oo.created_at as TIMESTAMP) < cast(concat(CURRENT_DATE, ' 23:59:59') as TIMESTAMP)
            ) as total,
            COALESCE((
                select sum(oi.quantity)
                from orders_orderitem oi
                join orders_order oo on oo.id = oi.order_id
                where oi.product_id = p.id and
                oi.product_id in (26, 27, 52) and
                oo.restaurant_branch_id in (113, 114,   214, 215, 216, 127, 108, 109, 111, 113,    10, 11, 12, 13, 14, 281,    4,5,6,7,8,236,     77, 235,331,     120, 212, 233, 279, 298, 310, 332, 366, 158, 397) and
                cast(oo.created_at as TIMESTAMP) > cast(concat(CURRENT_DATE, ' 00:00:00') as TIMESTAMP) and
                cast(oo.created_at as TIMESTAMP) < cast(concat(CURRENT_DATE, ' 23:59:59') as TIMESTAMP)
            ), 0) as sklad_french,
            COALESCE((
                select sum(oi.quantity)
                from orders_orderitem oi
                join orders_order oo on oo.id = oi.order_id
                where oi.product_id = p.id and
                oi.product_id in (161) and
                oo.restaurant_branch_id in (113, 114,   108, 109, 110, 111, 112, 127, 158, 215, 216, 214) and
                cast(oo.created_at as TIMESTAMP) > cast(concat(CURRENT_DATE, ' 00:00:00') as TIMESTAMP) and
                cast(oo.created_at as TIMESTAMP) < cast(concat(CURRENT_DATE, ' 23:59:59') as TIMESTAMP)
            ), 0) as sklad_brownie,
            pr.name as provider_name,
            pr.id as provider_id,
            p.id as product_id
            from products_product p
            join orders_orderitem oi on oi.product_id = p.id
            join products_productprovider pr on pr.id = p.provider_id
            join orders_order oo on oo.id = oi.order_id
            where
            cast(oo.created_at as TIMESTAMP) > cast(concat(CURRENT_DATE, ' 00:00:00') as TIMESTAMP) and
            cast(oo.created_at as TIMESTAMP) < cast(concat(CURRENT_DATE, ' 23:59:59') as TIMESTAMP) and
            p.id in (26, 27, 52, 161)
        `)
        .then(res => {
            if (res.rows) {
                const totalBrownie = res.rows.find(product => product.product_id == 161)?.total;
                const skladBrownie = res.rows.find(product => product.product_id == 161)?.sklad_brownie;
                const totalFrenchCheese = res.rows.find(product => product.product_id == 26)?.total;
                const skladFrenchCheese = res.rows.find(product => product.product_id == 26)?.sklad_french;
                const totalFrenchKetchup = res.rows.find(product => product.product_id == 27)?.total;
                const skladFrenchKetchup = res.rows.find(product => product.product_id == 27)?.sklad_french;
                const totalFrenchGorch = res.rows.find(product => product.product_id == 52)?.total;
                const skladFrenchGorch = res.rows.find(product => product.product_id == 52)?.sklad_french;
                let message = 'За сегодняшний период заявок:\n\n';
                message += `Общее кол-во Брауни - ${totalBrownie}. \nКате отправляем ${totalBrownie - skladBrownie}, \nа остальные ${skladBrownie} делаем сами.`;
                message += `\n\nОбщее кол-во Френчдогов с Сыром - ${totalFrenchCheese}. \nКате отправляем ${totalFrenchCheese - skladFrenchCheese}, \nа остальные ${skladFrenchCheese} делаем сами.`
                message += `\n\nОбщее кол-во Френчдогов с Кетчупом - ${totalFrenchKetchup}. \nКате отправляем ${totalFrenchKetchup - skladFrenchKetchup}, \nа остальные ${skladFrenchKetchup} делаем сами.`
                message += `\n\nОбщее кол-во Френчдогов с Горчицей - ${totalFrenchGorch}. \nКате отправляем ${totalFrenchGorch - skladFrenchGorch}, \nа остальные ${skladFrenchGorch} делаем сами.`
                bot.sendMessage(chatId, message);
            }
        })
      } else if (msg.text.toLowerCase().includes('маршрут')) {
        client.query(`
            select concat(b."name", ': ', b.address) as address
            from orders_order oo
            join profiles_restaurantbranch b on b.id = oo.restaurant_branch_id
            where
            oo.is_completed is true and
            cast(oo.created_at as TIMESTAMP) > cast(concat(CURRENT_DATE, ' 00:00:00') as TIMESTAMP) and
            cast(oo.created_at as TIMESTAMP) < cast(concat(CURRENT_DATE, ' 23:59:59') as TIMESTAMP)    
        `)
        .then(res => {
            const EXCELL_1_ROW = [
                {
                    value: 'Адрес',
                    fontWeight: 'bold',
                    width: 100,
                },
                {
                    value: 'Курьер',
                    width: 14,
                }
            ];

            const excellRows = [EXCELL_1_ROW];

            for (const item of res.rows) {
                excellRows.push([
                    {
                        value: item.address,
                        width: 50,
                    },
                    {
                        value: null,
                        width: 14,
                    }
                ])
            }

            const excellFilePath = path.join(__dirname, `маршруты-${moment().format('DD-MM-YYYY')}-${Math.round(Math.random()*100000)}.xlsx`);

            writeXlsxFile(
                excellRows,
                {
                    columns: [
                        { width: 50 },
                        { width: 15 },
                    ],
                    filePath: excellFilePath,
                },
            ).then(() => {
                const fileBuffer = fs.readFileSync(excellFilePath);

                bot.sendDocument(chatId, fileBuffer, {}, {
                    filename: `маршруты-${moment().format('DD-MM-YYYY')}-${Math.round(Math.random()*100000)}.xlsx`,
                    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                }); 
            })
        })
      } else if (msg.text.toLowerCase().includes('total daily')) {
        client.query(`
            select
            distinct p.name as product_name,
            (
                select sum(oi.quantity)
                from orders_orderitem oi
                join orders_order oo on oo.id = oi.order_id
                where oi.product_id = p.id and
                cast(oo.created_at as TIMESTAMP) > cast(concat(CURRENT_DATE, ' 00:00:00') as TIMESTAMP) and
                cast(oo.created_at as TIMESTAMP) < cast(concat(CURRENT_DATE, ' 23:59:59') as TIMESTAMP)
            ) - COALESCE((
                select sum(oi.quantity)
                from orders_orderitem oi
                join orders_order oo on oo.id = oi.order_id
                where oi.product_id = p.id and
                oi.product_id in (26, 27, 52) and
                oo.restaurant_branch_id in (113, 114,    214, 215, 216, 127, 108, 109, 111, 113,     10, 11, 12, 13, 14, 281,    4,5,6,7,8,236,     77, 235,331,     120, 212, 233, 279, 298, 310, 332, 366) and
                cast(oo.created_at as TIMESTAMP) > cast(concat(CURRENT_DATE, ' 00:00:00') as TIMESTAMP) and
                cast(oo.created_at as TIMESTAMP) < cast(concat(CURRENT_DATE, ' 23:59:59') as TIMESTAMP)
            ), 0) - COALESCE((
                select sum(oi.quantity)
                from orders_orderitem oi
                join orders_order oo on oo.id = oi.order_id
                where oi.product_id = p.id and
                oi.product_id in (161) and
                oo.restaurant_branch_id in (113, 114,    108, 109, 110, 111, 112, 127, 158, 215, 216, 214) and
                cast(oo.created_at as TIMESTAMP) > cast(concat(CURRENT_DATE, ' 00:00:00') as TIMESTAMP) and
                cast(oo.created_at as TIMESTAMP) < cast(concat(CURRENT_DATE, ' 23:59:59') as TIMESTAMP)
            ), 0) as count,
            pr.name as provider_name,
            pr.id as provider_id,
            p.id as product_id
            from products_product p
            join orders_orderitem oi on oi.product_id = p.id
            join products_productprovider pr on pr.id = p.provider_id
            join orders_order oo on oo.id = oi.order_id
            where
            cast(oo.created_at as TIMESTAMP) > cast(concat(CURRENT_DATE, ' 00:00:00') as TIMESTAMP) and
            cast(oo.created_at as TIMESTAMP) < cast(concat(CURRENT_DATE, ' 23:59:59') as TIMESTAMP)
        `)
        .then(res => {
            let providers = [...new Set(res.rows.map(product => product.provider_id))];
            providers = providers.map(id => ({
                id: +id,
                name: res.rows.find(product => product.provider_id == id)?.provider_name,
            }));

            for (const provider of providers) {
                const providerProducts = res.rows.filter(product => product.provider_id == provider.id);
                let message = `Заявка на ${moment().add(1, 'days').format('DD/MM/YYYY')}.\n${provider.name}\n`;
                for (const product of providerProducts) {
                    message = message.concat(`\n${product.product_name} - ${product.count} шт.`);
                }

                bot.sendMessage(chatId, message);
            }

            console.log(`[${moment().format('DD-MM-YYYY hh:m:ss')}] Пользователю ${msg.chat.username} отправлен total daily`);
        })
      } else if (msg.text.toLowerCase().includes('nak ')) {
        const providerName = msg.text.split(' ')[1] || null;
        if (providerName) {
            client.query('select id, name from products_productprovider')
            .then(res => {
                if (res.rows.length) {
                    const provider = res.rows.find(provider => provider.name.toLowerCase().includes(providerName.toLowerCase()));
                    if (provider) {
                        bot.sendMessage(chatId, "Номер найденного поставщика: " + provider.id + ". Название: " + provider.name);
                        
                        client.query(`
                            select oo.order_number, au.username, p.id as product_id, p.name as product_name, oi.quantity, oi.price, rb.id as rest_branch_id, rb.name as rest_name, rb.address as rest_address, rb.contact_phone, (
                                select sum(oi.quantity) from orders_orderitem oi
                                join products_product p on p.id = oi.product_id
                                where p.provider_id = ${provider.id} and oi.order_id = oo.id
                            ) as total_food_count, oo.created_at
                            from orders_order oo
                            join profiles_restaurantbranch rb on rb.id = oo.restaurant_branch_id
                            join orders_orderitem oi on oi.order_id = oo.id
                            join products_product p on p.id = oi.product_id
                            join auth_user au on au.id = oo.user_id
                            where
                            p.provider_id = ${provider.id} and
                            cast(oo.created_at as TIMESTAMP) > cast(concat(CURRENT_DATE, ' 00:00:00') as TIMESTAMP) and
                            cast(oo.created_at as TIMESTAMP) < cast(concat(CURRENT_DATE, ' 21:00:00') as TIMESTAMP) and
                            (
                                select count(oi.id) from orders_orderitem oi
                                join products_product p on p.id = oi.product_id
                                where p.provider_id = ${provider.id} and oi.order_id = oo.id
                            ) > 0
                        `)
                        .then(res => {
                            const products = res.rows;
                            if (products.length) {
                                const restaurantBranchIds = [...new Set(products.map(product => product?.rest_branch_id))];

                                const EXCELL_TOTAL_SHEETS = [];
                                const EXCELL_TOTAL_SHEETS_NAMES = [];
                                const EXCELL_TOTAL_COLUMNS = [];
                                
                                for (const branchId of restaurantBranchIds) {
                                    const branchProducts = products.filter(product => product.rest_branch_id === branchId);

                                    const EXCELL_1_ROW = [
                                        {
                                            value: 'Номер заказа',
                                            fontWeight: 'bold',
                                            width: 100,
                                        },
                                        {
                                            value: branchProducts[0].order_number,
                                            width: 10,
                                        }
                                    ];

                                    const EXCELL_2_ROW = [
                                        {
                                            value: 'Пользователь',
                                            fontWeight: 'bold',
                                            width: 15,
                                        },
                                        {
                                            value: branchProducts[0].username,
                                            width: 30,
                                        }
                                    ];

                                    const EXCELL_3_ROW = [
                                        {
                                            value: 'Ресторан',
                                            fontWeight: 'bold',
                                            width: 60,
                                        },
                                        {
                                            value: branchProducts[0].rest_name + ' - ' + branchProducts[0].rest_address,
                                            width: 60,
                                        }
                                    ];

                                    const EXCELL_4_ROW = [
                                        {
                                            value: 'Дата создания',
                                            fontWeight: 'bold',
                                            width: 20,
                                        },
                                        {
                                            value: moment(branchProducts[0].created_at).add(1, 'days').format("YYYY-MM-DD"),
                                            width: 15,
                                        }
                                    ];

                                    const EXCELL_5_ROW = [
                                        {
                                            value: '#',
                                            fontWeight: 'bold',
                                        },
                                        {
                                            value: 'Наименование',
                                            fontWeight: 'bold',
                                            width: 300,
                                        },
                                        {
                                            value: 'Количество',
                                            fontWeight: 'bold',
                                            width: 30,
                                        },
                                        {
                                            value: 'Цена',
                                            fontWeight: 'bold',
                                            width: 15,
                                        },
                                        {
                                            value: 'Сумма',
                                            fontWeight: 'bold',
                                            width: 15,
                                        },
                                    ];

                                    const EXCELL_ROWS_ARRAY = [EXCELL_1_ROW, EXCELL_2_ROW, EXCELL_3_ROW, EXCELL_4_ROW, EXCELL_5_ROW];

                                    let totalBranchProductsCount = 0;
                                    let totalBranchProductsPrice = 0;
                                    let orderNumber = 1;

                                    // console.log(branchProducts)

                                    for (const product of branchProducts) {
                                        if (
                                            ([113, 114,   214, 215, 216, 127, 108, 109, 111, 113, 10, 11, 12, 13, 14, 281, 4,5,6,7,8,236, 77, 235,331, 120, 212, 233, 279, 298, 310, 332, 366].includes(+product.rest_branch_id) &&
                                            [26, 27, 52].includes(+product.product_id)) ||
                                            ((product.product_id == 161) &&
                                            ([113, 114,   108, 109, 110, 111, 112, 127, 158, 215, 216, 214].includes(+product.rest_branch_id)))
                                        ) {
                                            
                                        } else {
                                            EXCELL_ROWS_ARRAY.push([
                                                {
                                                    value: orderNumber,
                                                },
                                                {
                                                    value: product.product_name,
                                                },
                                                {
                                                    value: product.quantity,
                                                },
                                                {
                                                    value: product.price,
                                                },
                                                {
                                                    value: product.price * product.quantity,
                                                },
                                            ]);
    
                                            totalBranchProductsCount += product.quantity;
                                            totalBranchProductsPrice += product.price * product.quantity;
                                            orderNumber++;
                                        }
                                    }

                                    EXCELL_ROWS_ARRAY.push([
                                        {
                                            value: null,
                                        },
                                        {
                                            value: 'Итого',
                                        },
                                        {
                                            value: totalBranchProductsCount,
                                        },
                                        {
                                            value: null,
                                        },
                                        {
                                            value: totalBranchProductsPrice,
                                        },
                                    ]);

                                    EXCELL_TOTAL_SHEETS_NAMES.push(branchProducts[0].rest_address.substring(0, 29).replaceAll('/', '-'));
                                    EXCELL_TOTAL_SHEETS.push(EXCELL_ROWS_ARRAY);
                                    EXCELL_TOTAL_COLUMNS.push([
                                        { width: 15 },
                                        { width: 50 },
                                        { width: 12 },
                                        { width: 10 },
                                        { width: 10 },
                                    ]);
                                }

                                const excellFilePath = path.join(__dirname, `Накладные-${provider.name.replaceAll(':', ' ').replaceAll(' ', '_').replaceAll('"', '')}-${moment().add(1, 'days').format("YYYY_MM_DD")}-${Math.round(Math.random()*100000)}.xlsx`);
                                fs.writeFile(excellFilePath, '', () => {
                                    writeXlsxFile(
                                        EXCELL_TOTAL_SHEETS,
                                        {
                                            sheets: EXCELL_TOTAL_SHEETS_NAMES,
                                            columns: EXCELL_TOTAL_COLUMNS,
                                            filePath: excellFilePath,
                                        },
                                    )
                                    .then(() => {
                                        const fileBuffer = fs.readFileSync(excellFilePath);
    
                                        bot.sendDocument(chatId, fileBuffer, {}, {
                                            filename: `Накладные-${provider.name.replaceAll('Поставщик', '').replaceAll(':', ' ').replaceAll(' ', '_').replaceAll('"', '')}-${moment().add(1, 'days').format("YYYY_MM_DD")}-${Math.round(Math.random()*100000)}.xlsx`,
                                            contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                        }); 
                                    })
                                })
                            }
                        })

                    } else {
                        bot.sendMessage(chatId, "Поставщик не найден, возможно неправильно введено название");
                    }
                } else {
                    bot.sendMessage(chatId, "Поставщик не найден, возможно неправильно введено название");
                }
            })
        } else {
            bot.sendMessage(chatId, 'Неверный формат сообщения, нужно написать в формате "nak [название поставщика]"');
        }
      } else if (msg.text.toLowerCase().includes('наклад')) {
        client.query('select id, name from products_productprovider')
            .then(res => {
                client.query(`
                    select oo.order_number, au.username, p.id as product_id, p.name as product_name, oi.quantity, oi.price, rb.id as rest_branch_id, rb.name as rest_name, rb.address as rest_address, rb.contact_phone, p.provider_id, prov.name as provider_name,
                    (
                        select sum(oi.quantity) from orders_orderitem oi
                        join products_product p on p.id = oi.product_id
                        where p.provider_id = p.provider_id and oi.order_id = oo.id
                    ) as total_food_count, oo.created_at
                    from orders_order oo
                    join profiles_restaurantbranch rb on rb.id = oo.restaurant_branch_id
                    join orders_orderitem oi on oi.order_id = oo.id
                    join products_product p on p.id = oi.product_id
                    join auth_user au on au.id = oo.user_id
                    join products_productprovider prov on p.provider_id = prov.id
                    where
                    cast(oo.created_at as TIMESTAMP) > cast(concat(CURRENT_DATE, ' 00:00:00') as TIMESTAMP) and
                    cast(oo.created_at as TIMESTAMP) < cast(concat(CURRENT_DATE, ' 21:00:00') as TIMESTAMP) and
                    (
                        select count(oi.id) from orders_orderitem oi
                        join products_product p on p.id = oi.product_id
                        where p.provider_id = p.provider_id and oi.order_id = oo.id
                    ) > 0
                `)
                .then(res => {
                    
                    const allProducts = res.rows;
                    const providerIds = [...new Set(allProducts.map(item => item.provider_id))];

                    for (let providerId of providerIds) {
                        const products = allProducts.filter(product => product.provider_id == providerId);
                        if (products.length) {
                            const restaurantBranchIds = [...new Set(products.map(product => product?.rest_branch_id))];

                            const EXCELL_TOTAL_SHEETS = [];
                            const EXCELL_TOTAL_SHEETS_NAMES = [];
                            const EXCELL_TOTAL_COLUMNS = [];
                            
                            for (const branchId of restaurantBranchIds) {
                                const branchProducts = products.filter(product => product.rest_branch_id === branchId);

                                const EXCELL_1_ROW = [
                                    {
                                        value: 'Номер заказа',
                                        fontWeight: 'bold',
                                        width: 100,
                                    },
                                    {
                                        value: branchProducts[0].order_number,
                                        width: 10,
                                    }
                                ];

                                const EXCELL_2_ROW = [
                                    {
                                        value: 'Пользователь',
                                        fontWeight: 'bold',
                                        width: 15,
                                    },
                                    {
                                        value: branchProducts[0].username,
                                        width: 30,
                                    }
                                ];

                                const EXCELL_3_ROW = [
                                    {
                                        value: 'Ресторан',
                                        fontWeight: 'bold',
                                        width: 60,
                                    },
                                    {
                                        value: branchProducts[0].rest_name + ' - ' + branchProducts[0].rest_address,
                                        width: 60,
                                    }
                                ];

                                const EXCELL_4_ROW = [
                                    {
                                        value: 'Дата создания',
                                        fontWeight: 'bold',
                                        width: 20,
                                    },
                                    {
                                        value: moment(branchProducts[0].created_at).add(1, 'days').format("YYYY-MM-DD"),
                                        width: 15,
                                    }
                                ];

                                const EXCELL_5_ROW = [
                                    {
                                        value: '#',
                                        fontWeight: 'bold',
                                    },
                                    {
                                        value: 'Наименование',
                                        fontWeight: 'bold',
                                        width: 300,
                                    },
                                    {
                                        value: 'Количество',
                                        fontWeight: 'bold',
                                        width: 30,
                                    },
                                    {
                                        value: 'Цена',
                                        fontWeight: 'bold',
                                        width: 15,
                                    },
                                    {
                                        value: 'Сумма',
                                        fontWeight: 'bold',
                                        width: 15,
                                    },
                                ];

                                const EXCELL_ROWS_ARRAY = [EXCELL_1_ROW, EXCELL_2_ROW, EXCELL_3_ROW, EXCELL_4_ROW, EXCELL_5_ROW];

                                let totalBranchProductsCount = 0;
                                let totalBranchProductsPrice = 0;
                                let orderNumber = 1;

                                // console.log(branchProducts)

                                for (const product of branchProducts) {
                                    if (
                                        ([113, 114,   214, 215, 216, 127, 108, 109, 111, 113, 10, 11, 12, 13, 14, 281, 4,5,6,7,8,236, 77, 235,331, 120, 212, 233, 279, 298, 310, 332, 366].includes(+product.rest_branch_id) &&
                                        [26, 27, 52].includes(+product.product_id)) ||
                                        ((product.product_id == 161) &&
                                        ([113, 114,   108, 109, 110, 111, 112, 127, 158, 215, 216, 214].includes(+product.rest_branch_id)))
                                    ) {
                                        
                                    } else {
                                        EXCELL_ROWS_ARRAY.push([
                                            {
                                                value: orderNumber,
                                            },
                                            {
                                                value: product.product_name,
                                            },
                                            {
                                                value: product.quantity,
                                            },
                                            {
                                                value: product.price,
                                            },
                                            {
                                                value: product.price * product.quantity,
                                            },
                                        ]);

                                        totalBranchProductsCount += product.quantity;
                                        totalBranchProductsPrice += product.price * product.quantity;
                                        orderNumber++;
                                    }
                                }

                                EXCELL_ROWS_ARRAY.push([
                                    {
                                        value: null,
                                    },
                                    {
                                        value: 'Итого',
                                    },
                                    {
                                        value: totalBranchProductsCount,
                                    },
                                    {
                                        value: null,
                                    },
                                    {
                                        value: totalBranchProductsPrice,
                                    },
                                ]);

                                EXCELL_TOTAL_SHEETS_NAMES.push(branchProducts[0].rest_address.substring(0, 29).replaceAll('/', '-'));
                                EXCELL_TOTAL_SHEETS.push(EXCELL_ROWS_ARRAY);
                                EXCELL_TOTAL_COLUMNS.push([
                                    { width: 15 },
                                    { width: 50 },
                                    { width: 12 },
                                    { width: 10 },
                                    { width: 10 },
                                ]);
                            }

                            const excellFilePath = path.join(__dirname, `Накладные-${products[0].provider_name.replaceAll(':', ' ').replaceAll(' ', '_').replaceAll('"', '')}-${moment().add(1, 'days').format("YYYY_MM_DD")}-${Math.round(Math.random()*100000)}.xlsx`);
                            fs.writeFile(excellFilePath, '', () => {
                                writeXlsxFile(
                                    EXCELL_TOTAL_SHEETS,
                                    {
                                        sheets: EXCELL_TOTAL_SHEETS_NAMES,
                                        columns: EXCELL_TOTAL_COLUMNS,
                                        filePath: excellFilePath,
                                    },
                                )
                                .then(() => {
                                    const fileBuffer = fs.readFileSync(excellFilePath);

                                    bot.sendDocument(chatId, fileBuffer, {}, {
                                        filename: `Накладные-${products[0].provider_name.replaceAll('Поставщик', '').replaceAll(':', ' ').replaceAll(' ', '_').replaceAll('"', '')}-${moment().add(1, 'days').format("YYYY_MM_DD")}-${Math.round(Math.random()*100000)}.xlsx`,
                                        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                    }); 
                                })
                            })
                        }
                    }
                })
            })
      }
      else if (msg.text.split(' ').length) {
        const username = msg.text.split(' ')[1];
        bot.sendMessage(chatId, "Секунду...");
  
        client.query(`
          select
          distinct o.user_id,
          u.username as account,
          SUBSTRING(CAST(u.date_joined as TEXT), 0, 11) as registration_date,
          (
              select count(o2.id)
              from orders_order o2
              where o2.user_id = o.user_id
          ) as orders_count,
          (
              select SUBSTRING(CAST(min(o4.created_at) as TEXT), 0, 11)
              from orders_order o4
              where o4.user_id = o.user_id
          ) as first_order_date,
          (
              select SUBSTRING(CAST(max(o4.created_at) as TEXT), 0, 11)
              from orders_order o4
              where o4.user_id = o.user_id
          ) as last_order_date,
          (
              select CAST(avg(o3.total_price) as INTEGER)
              from orders_order o3
              where o3.user_id = o.user_id
          ) as average_price,
          (
              select sum(o3.total_price)
              from orders_order o3
              where o3.user_id = o.user_id
          ) as total_sum
          from orders_order o
          join auth_user u on u.id = o.user_id
          where u.username='${username}' and o.is_completed is TRUE
        `)
        .then(res => {
          let finalMessage = null;
          if (res.rows.length) {
              finalMessage = `Информация по аккаунту ${username}:
Дата регистрации: ${res.rows[0].registration_date}
Кол-во заказов: ${res.rows[0].orders_count}
Дата первого заказа: ${res.rows[0].first_order_date}
Дата крайнего заказа: ${res.rows[0].last_order_date}
Средний чек: ${res.rows[0].average_price}
Общая сумма заказов: ${res.rows[0].total_sum}
`;
          } else {
              finalMessage = 'Этот аккаунт еще ничего не заказывал'
          }
  
          bot.sendMessage(chatId, finalMessage);
        }) 
      }
    });
})