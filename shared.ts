namespace MJS {


    export function throwf(err: Error): never {
        throw err;
    }


    export function sleep(time: number): Promise<{}> {
        return new Promise((resolve) => setTimeout(resolve, time));
    }


    export type Page_Data = {
        page:               string;
        browser_source:     string;
        page_window:        Page_Window;
        mdl_course_categories: Partial<MDL_Course_Categories>;
        mdl_course:         Partial<MDL_Course>;
        mdl_course_sections: Partial<MDL_Course_Sections>;
        mdl_course_modules: Partial<MDL_Course_Modules>;
        dom_submit:         boolean|string;
        dom_set_key:        string;
        dom_set_value:      boolean|number|string;
        error:              {message: string};
    }


    export type Page_Window = {
        location_origin:    string;
        sesskey:            string;
        body_id:            string;
        body_class:         string;
    }


    export type MDL_Context_Instance = {
        readonly id:        number;     // key
    }


    export type MDL_Course_Categories = MDL_Context_Instance & {
        // context level 40
        name:               string;
        idnumber:           string;
        description:        string;
        descriptionformat:  number;
        parent:             number;
        sortorder:          number;
        coursecount:        number;
        visible:            number;
        visibleold:         number;
        timemodified:       number;
        depth:              number;
        path:               string;
        theme:              string;
    }

    export type MDL_Course = MDL_Context_Instance & {
        // context level 50
        category:           number;     // -> course_categories.id
        sortorder:          number;
        fullname:           string;
        shortname:          string;
        idnumber:           string;
        summary:            string;
        //summaryformat:      number;
        format:             string;
        showgrades:         number;
        newsitems:          number;
        startdate:          number; // date?
        //enddate:            number;
        marker:             number;
        maxbytes:           number;
        //legacyfiles:        number;
        showreports:        number;
        visible:            number;
        //visibleold:         number;
        groupmode:          number;
        groupmodeforce:     number;
        defaultgroupingid:  number;
        lang:               string;
        //calendartype:       string;
        //theme
        //timecreated
        //timemodified
        //requested
        enablecompletion:   number;
        //completionnotify
        //cacherev
        mdl_course_sections: Partial<MDL_Course_Sections>[];
    }


    export type MDL_Course_Sections = {
        readonly id:        number;     // key
        course:             number;     // -> course.id
        section:            number;
        name:               string;
        summary:            string;
        summaryformat:      number;
        sequence:           string;
        visible:            number;
        availability:       string;
        timemodified:       number;
        x_options: {level?: number};
        //x_submit:           boolean;
        mdl_course_modules: Partial<MDL_Course_Modules>[];
    }


    // course_format_options

    
    export type MDL_Course_Modules = MDL_Context_Instance & {
        // context level 70
        course:             number;
        module:             number;
        mdl_modules_name: string;
        instance:           number;     // -> activity.id for type module
        mdl_course_module_instance: Partial<MDL_Course_Module_Instance>;
        section:            number;
        idnumber:           string|null;
        added:              number;
        score:              number;
        indent:             number;
        visible:            number;
        visibleoncoursepage: number;
        visibleold:         number;
        groupmode:          number;
        groupingid:         number;
        completion:         number;
        completiongradeitemnumber: number;
        completionview:     number;
        completionexpected: number;
        showdescription:    number;
        availability:       string;
        deletioninprogress: number;
        //x_submit:           boolean;
    }


    export type MDL_Course_Module_Instance_Abstract = {
        readonly id:        number;     // key
        course:             number;
        name:               string;
        intro:              string;
        introformat:        number;
    }


    export type MDL_Assignment = MDL_Course_Module_Instance_Abstract & {
        alwaysshowdescription: number;
        nosubmissions:      number;
        sendnotifications:  number;
        sendlatenotifications: number;
        duedate:            number;
        allowsubmissionsfromdate: number;
        grade:              number;
        timemodified:       number;
        requiresubmissionstatement: number;
        completionsubmit:   number;
        cutoffdate:         number;
        gradingduedate:     number;
        teamsubmission:     number;
        requireallteammemberssubmit: number;
        teamsubmissiongroupingid: number;
        blindmarking:       number;
        revealidentities:   number;
        attemptreopenmethod: string;
        maxattempts:        number;
        markingworkflow:    number;
        markingallocation:  number;
        sendstudentnotifications: number;
        preventsubmissionnotingroup: number;
    }


    export type MDL_Forum = MDL_Course_Module_Instance_Abstract & {
        type:               string;
        assessed:           number;
        assesstimestart:    number;
        assesstimefinish:   number;
        scale:              number;
        maxbytes:           number;
        maxattachments:     number;
        forcesubscribe:     number;
        trackingtype:       number;
        rsstype:            number;
        rssarticles:        number;
        timemodified:       number;
        warnafter:          number;
        blockafter:         number;
        blockperiod:        number;
        completiondiscussions: number;
        completionreplies:  number;
        completionposts:    number;
        displaywordcount:   number;
        lockdiscussionafter: number;
    }


    export type MDL_Feedback = MDL_Course_Module_Instance_Abstract & {
        mdl_feedback_template_id: number;
    }


    export type MDL_Course_Module_Instance = MDL_Assignment | MDL_Forum | MDL_Feedback;


}