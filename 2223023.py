def run(path):
    import pandas as pd
    import re

    attendance = pd.read_excel(path, sheet_name="attendance")
    students = pd.read_excel(path, sheet_name="students")

    attendance['attendance_date'] = pd.to_datetime(attendance['attendance_date'])

    attendance = attendance.sort_values(['student_id', 'attendance_date'])

    results = []

    for student_id, group in attendance.groupby('student_id'):
        group = group.reset_index(drop=True)
        group['is_absent'] = group['status'] == 'Absent'
        group['gap'] = (group['is_absent'] != group['is_absent'].shift()).cumsum()
        
        streaks = group[group['is_absent']].groupby('gap')
        
        valid_streaks = []

        for _, streak in streaks:
            if len(streak) >= 3:
                valid_streaks.append({
                    'student_id': student_id,
                    'absence_start_date': streak['attendance_date'].iloc[0],
                    'absence_end_date': streak['attendance_date'].iloc[-1],
                    'total_absent_days': len(streak)
                })

        if valid_streaks:
            latest = max(valid_streaks, key=lambda x: x['absence_end_date'])
            results.append(latest)

    df_absent = pd.DataFrame(results)

    if df_absent.empty:
        return pd.DataFrame(columns=['student_id', 'absence_start_date', 'absence_end_date', 'total_absent_days', 'email', 'msg'])

    df_final = df_absent.merge(students, on='student_id', how='left')

    def is_valid(email):
        pattern = r'^[A-Za-z_][A-Za-z0-9_]*@[\w]+\.(com)$'
        return bool(re.fullmatch(pattern, email))

    df_final['email'] = df_final['parent_email'].apply(lambda x: x if is_valid(str(x)) else '')

    def create_msg(row):
        if row['email'] != '':
            return f"Dear Parent, your child {row['student_name']} was absent from {row['absence_start_date'].date()} to {row['absence_end_date'].date()} for {row['total_absent_days']} days. Please ensure their attendance improves."
        else:
            return ''

    df_final['msg'] = df_final.apply(create_msg, axis=1)

    df_final = df_final[['student_id', 'absence_start_date', 'absence_end_date', 'total_absent_days', 'email', 'msg']]

    return df_final