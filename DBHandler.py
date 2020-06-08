class Read:

    def __init__(self):
        pass

    # Last Date
    def last_date(self, conn, gateway_db):  
        cursor = conn.cursor()  
        cursor.execute(f"SELECT max(deliveryDate) FROM {gateway_db}.[dbo].[OfflineMessage] where delivered=1")
        for row in cursor:
            date = str(row[0])
        return date

    # Current Sales
    def curr_sales(self, conn, gateway_db): 
        cursor = conn.cursor()   
        cursor.execute(f"SELECT Count(id) FROM {gateway_db}.dbo.OfflineMessage where delivered=1")
        for row in cursor:
            sales = row[0]
        return sales
    
    #Prev Sales
    def prev_sales(self, conn, gateway_db, deliveryDate):
        cursor = conn.cursor()
        cursor.execute(f"SELECT Count(id) FROM {gateway_db}.dbo.OfflineMessage where deliveryDate<='{deliveryDate}'")
        for row in cursor:
            sales = row[0]
        return sales


    def weekly_sales(self, conn_local, gateway_name):
        cursor = conn_local.cursor()
        cursor.execute(f"SELECT max(prevSales), max(currSales), max(incBy), max(lastDate) FROM WeeklySales.dbo.Sales where gateways = '{gateway_name}'")
        for row in cursor:
            sales = row
        return sales
        
class Write:
    
    def __init__(self):
        pass

    def write(self, conn, gateway_name, prev_sales, curr_sales, last_date):
        cursor = conn.cursor()
        cursor.execute(f'''
                    INSERT INTO WeeklySales.dbo.Sales (gateWays, prevSales
                    ,currSales, incBy, lastDate)
                    VALUES ('{gateway_name}', '{prev_sales}'
                    ,'{curr_sales}', {curr_sales - prev_sales}, '{last_date}')''')
        conn.commit()  