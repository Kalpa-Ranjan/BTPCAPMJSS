service SchedulerService @(path: '/api') @(requires: 'any') {

    action scheduleJob(payload: LargeString) returns LargeString;
    
}