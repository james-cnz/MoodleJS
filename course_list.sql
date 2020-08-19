SELECT cc.path as "Cat Path", c.id as "Course ID", c.shortname as "Course Short Name",
    CONCAT('%%WWWROOT%%/course/view.php',
      '%%Q%%id=', c.id) as "Course URL"
FROM prefix_course AS c
LEFT JOIN prefix_course_categories AS cc
    on cc.id = c.category
ORDER by cc.path, c.shortname