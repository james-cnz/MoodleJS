/*
 * Moodle JS Misc Routines
 * Used in background and content scripts.
 */


namespace MJS {

    export type DeepPartial<T> = {
        [P in keyof T]?: T[P] extends Array<infer U> // tslint:disable-line:array-type
          ? Array<DeepPartial<U>> // tslint:disable-line:array-type
          : T[P] extends ReadonlyArray<infer U> // tslint:disable-line:no-shadowed-variable
            ? ReadonlyArray<DeepPartial<U>>
            : DeepPartial<T[P]>
      };

    // tslint:disable-next-line: promise-function-async
    export function sleep(time: number): Promise<{}> {
        return new Promise((resolve) => setTimeout(resolve, time));
    }

    export function throwf(err: Error): never {
        throw err;
    }

    export type Errorlike = {
        name:           string;
        message:        string;
        fileName?:      string;
        lineNumber?:    number;
    };

    export function is_Errorlike(possible_Errorlike: any): possible_Errorlike is Errorlike {
        return (typeof possible_Errorlike == "object")
            && (possible_Errorlike != null)
            && (typeof ((possible_Errorlike as Errorlike).name) == "string")
            && (typeof ((possible_Errorlike as Errorlike).message) == "string");
    }





    export type MDL_Context = {
        readonly context_id:    number;     // key
        readonly contextlevel:  number;
    };


    export type MDL_Course_Category = MDL_Context & {
        readonly course_category_id: number;
        readonly contextlevel: 40;
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
        mdl_course_categories: MDL_Course_Category[];
        mdl_courses:        MDL_Course[];
    };

    export type MDL_Course = MDL_Context & {
        readonly course_id: number;
        readonly contextlevel: 50;
        category:           number;     // -> course_categories.id
        sortorder:          number;
        fullname:           string;
        shortname:          string;
        idnumber:           string;
        summary:            string;
        // summaryformat:      number;
        format:             string;
        showgrades:         number;
        newsitems:          number;
        startdate:          number; // date?
        // enddate:            number;
        marker:             number;
        maxbytes:           number;
        // legacyfiles:        number;
        showreports:        number;
        visible:            number;
        // visibleold:         number;
        groupmode:          number;
        groupmodeforce:     number;
        defaultgroupingid:  number;
        lang:               string;
        // calendartype:       string;
        // theme
        // timecreated
        // timemodified
        // requested
        enablecompletion:   number;
        // completionnotify
        // cacherev
        mdl_course_sections: MDL_Course_Section[];
    };


    export type MDL_Course_Section = {
        readonly course_section_id: number;
        course:             number;     // -> course.id
        section:            number;
        name:               string;
        summary:            string;
        summaryformat:      number;
        sequence:           string;
        visible:            number;
        availability:       string;
        timemodified:       number;
        options:            {level?: number};
        mdl_course_modules: MDL_Course_Module[];
    };


    // course_format_options


    export type MDL_Course_Module_Base = MDL_Context & {
        readonly course_module_id: number;
        readonly contextlevel: 70;
        course:             number;
        module:             number;
        mdl_module_name:    string;
        // instance:           number;     // -> activity.id for type module
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

        readonly activity_id: number;
        // course:             number;
        name:               string;
        intro:              string;
        introformat:        number;
    };


    export type MDL_Assignment = MDL_Course_Module_Base & {
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
    };


    export type MDL_Forum = MDL_Course_Module_Base & {
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
    };


    export type MDL_Feedback = MDL_Course_Module_Base & {
        mdl_feedback_template_id: number;
    };


    export type MDL_Course_Module = MDL_Assignment | MDL_Forum | MDL_Feedback;

    export type MDL_User = MDL_Context & {
        readonly contextlevel: 30;
        username:           string;
        password:           string;
    };

}
