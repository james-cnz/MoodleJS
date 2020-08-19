SELECT cc.id as "Cat ID",

(select string_agg(name, ' / ') from
(with recursive cc_tree as (
   select id, parent, name, 1 as count
   from prefix_course_categories
   where id = cc.id
   union all
   select cc_p.id, cc_p.parent, cc_p.name, (cc_c.count + 1) as count from
   cc_tree cc_c
   join
   prefix_course_categories cc_p on cc_c.parent = cc_p.id
) 
select name
from cc_tree
order by count desc) as cc_tree_ordered)

as "Cat Path",
    c.id as "Course ID", c.shortname as "Course Name",
    CONCAT('%%WWWROOT%%/course/view.php',
      '%%Q%%id=', c.id) as "URL"
FROM prefix_course AS c
LEFT JOIN prefix_course_categories AS cc
    on cc.id = c.category
ORDER by "Cat Path", c.shortname