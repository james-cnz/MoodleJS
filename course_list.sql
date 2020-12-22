SELECT cc.path as "Cat Path", c.id as "Course ID", c.shortname as "Course Short Name",
    CONCAT('%%WWWROOT%%/course/view.php', '%%Q%%id=', c.id) as "Course URL",
    lsl_max.max_timecreated as "Last Modified"
FROM prefix_course AS c
LEFT JOIN prefix_course_categories AS cc
    on cc.id = c.category
LEFT JOIN
    (SELECT courseid, MAX(timecreated) AS max_timecreated
     FROM prefix_logstore_standard_log
     WHERE crud <> 'r'
       AND target <> 'course_backup'
       AND userid NOT IN (
           SELECT id FROM prefix_user
           WHERE firstname = 'Link checker' AND lastname = 'Robot'
       ) /*76141*/
     GROUP BY courseid) AS lsl_max
    on lsl_max.courseid = c.id
WHERE lsl_max.max_timecreated >= 0 /*modified_since*/
ORDER by cc.path, c.shortname