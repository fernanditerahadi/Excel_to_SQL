--1 Write an SQL statement to count the number of users per country (5 marks)
SELECT country, COUNT(DISTINCT userid) AS NumberOfUsers
                    FROM user_tab
                    GROUP BY country;

--2 Write an SQL statement to count the number of orders per country (10 marks)
SELECT user_tab.country AS Country, COUNT(orderid) AS NumberOfOrders
                    FROM user_tab JOIN order_tab
                    ON user_tab.userid=order_tab.userid
                    GROUP BY user_tab.country

--3 Write an SQL statement to find the first order date of each user (10 marks)
SELECT userid, MIN(order_time) AS FirstOrderTime
        FROM order_tab
        GROUP BY userid

--4 Write an SQL statement to find the number of users who made their first order in each country, each day (25 marks)
SELECT u.country AS Country, COUNT(u.userid) AS NumOfUsers
                    FROM user_tab u JOIN order_tab o ON u.userid = o.userid
                    WHERE o.order_time IN (SELECT MIN(order_time)
                                            FROM order_tab o JOIN user_tab u
                                            ON o.userid=u.userid GROUP BY u.country)
                    GROUP BY u.country
--5 Write an SQL statement to find the first order GMV of each user. If there is a tie, use the order with the lower orderid (30 marks)
SELECT x.userid AS UserID, x.gmv AS GMV, x.orderid AS OrderID
FROM (SELECT a.userid, a.gmv, a.orderid, COUNT(*) AS FirstOrderGMV
      FROM order_tab a JOIN order_tab b
      ON a.userid = b.userid
      WHERE a.gmv < b.gmv OR (a.gmv = b.gmv AND a.orderid >= b.orderid)
      GROUP BY a.userid, a.gmv, a.orderid
      HAVING FirstOrderGMV = 1) AS x

--6 Find out what is wrong with the sample data (20 marks)
1. Some values in order_time column are earlier than the values found in register_time column.
2. There are duplicates in country column in user_tab table.
3. There should be another table that assigns country's id to each corresponding country's name
