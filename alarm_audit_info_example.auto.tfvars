####################
### Xpath: /managed-element/container-resource-management/audit-management/alarm-audit-info
### CLI Commands: CHG-AUD-INF:[STS],[PRD],[DATE],[DAY],[TIME];
### Description:
### This list indicates periodic audit information configured in the system.
### This operation is performed with target defined in the item.
### Audit interval can be specified by day, date and time.
### Dependencies: None
####################

alarm_audit_info = [
  {
    alarm_audit_item = "alarm"
    alarm_audit_status = "alw"
    alarm_audit_period = "hourly"
    alarm_audit_time = "40"
  },
]
