from random import shuffle

a = [{'Торговый зал АСФ №73': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №73.xlsx', 'АСФ №73', '73']}, {'Торговый зал АФ №63': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №63.xlsx', 'АФ №63', '63']}, {'Торговый зал ШФ №26': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №26.xlsx', 'ШФ №26', '26']}, {'Торговый зал АФ №26': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №26.xlsx', 'АФ №26', '26']}, {'Торговый зал АСФ №38': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №38.xlsx', 'АСФ №38', '38']}, {'Торговый зал АСФ №39': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №39.xlsx', 'АСФ №39', '39']}, {'Торговый зал ШФ №12': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №12.xlsx', 'ШФ №12', '12']}, {'Торговый зал ЕКФ №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ЕКФ №1.xlsx', 'ЕКФ №1', '1']}, {'Торговый зал АФ №83': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №83.xlsx', 'АФ №83', '83']}, {'Торговый зал ППФ №3': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №3.xlsx', 'ППФ №3', '3']}, {'Торговый зал АФ №31': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №31.xlsx', 'АФ №31', '31']}, {'Торговый зал АФ №7': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №7.xlsx', 'АФ №7', '7']}, {'Торговый зал ШФ №33': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №33.xlsx', 'ШФ №33', '33']}, {'Торговый зал АСФ №79': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №79.xlsx', 'АСФ №79', '79']}, {'Торговый зал АСФ №47': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №47.xlsx', 'АСФ №47', '47']}, {'Торговый зал АСФ №11': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №11.xlsx', 'АСФ №11', '11']}, {'Торговый зал АСФ №14': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №14.xlsx', 'АСФ №14', '14']}, {'Торговый зал ТЗФ №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ТЗФ №1.xlsx', 'ТЗФ №1', '1']}, {'Торговый зал АСФ №27': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №27.xlsx', 'АСФ №27', '27']}, {'Торговый зал ТФ №4': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ТФ №4.xlsx', 'ТФ №4', '4']}, {'Торговый зал АСФ №19': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №19.xlsx', 'АСФ №19', '19']}, {'Торговый зал УКФ №3': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал УКФ №3.xlsx', 'УКФ №3', '3']}, {'Торговый зал АФ №53': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №53.xlsx', 'АФ №53', '53']}, {'Торговый зал АФ №22': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №22.xlsx', 'АФ №22', '22']}, {'Торговый зал АФ №5': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №5.xlsx', 'АФ №5', '5']}, {'Торговый зал АСФ №32': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №32.xlsx', 'АСФ №32', '32']}, {'Торговый зал АСФ №81': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №81.xlsx', 'АСФ №81', '81']}, {'Торговый зал АСФ №48': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №48.xlsx', 'АСФ №48', '48']}, {'Торговый зал АФ №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №2.xlsx', 'АФ №2', '2']}, {'Торговый зал АФ №42': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №42.xlsx', 'АФ №42', '42']}, {'Торговый зал АСФ №26': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №26.xlsx', 'АСФ №26', '26']}, {'Торговый зал КПФ №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал КПФ №1.xlsx', 'КПФ №1', '1']}, {'Торговый зал АФ №66': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №66.xlsx', 'АФ №66', '66']}, {'Торговый зал ШФ №23': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №23.xlsx', 'ШФ №23', '23']}, {'Торговый зал ШФ №13': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №13.xlsx', 'ШФ №13', '13']}, {'Торговый зал АФ №3': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №3.xlsx', 'АФ №3', '3']}, {'Торговый зал АФ №60': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №60.xlsx', 'АФ №60', '60']}, {'Торговый зал АФ №40': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №40.xlsx', 'АФ №40', '40']}, {'Торговый зал АФ №67': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №67.xlsx', 'АФ №67', '67']}, {'Торговый зал АФ №73': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №73.xlsx', 'АФ №73', '73']}, {'Торговый зал АФ №69': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №69.xlsx', 'АФ №69', '69']}, {'Торговый зал ШФ №28': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №28.xlsx', 'ШФ №28', '28']}, {'Торговый зал АФ №6': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №6.xlsx', 'АФ №6', '6']}, {'Торговый зал АФ №11': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №11.xlsx', 'АФ №11', '11']}, {'Торговый зал АФ №82': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №82.xlsx', 'АФ №82', '82']}, {'Торговый зал АСФ №74': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №74.xlsx', 'АСФ №74', '74']}, {'Торговый зал ППФ №19': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №19.xlsx', 'ППФ №19', '19']}, {'Торговый зал ППФ №11': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №11.xlsx', 'ППФ №11', '11']}, {'Торговый зал ШФ №22': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №22.xlsx', 'ШФ №22', '22']}, {'Торговый зал АСФ №20': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №20.xlsx', 'АСФ №20', '20']}, {'Торговый зал АСФ №60': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №60.xlsx', 'АСФ №60', '60']}, {'Торговый зал ШФ №21': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №21.xlsx', 'ШФ №21', '21']}, {'Торговый зал АФ №35': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №35.xlsx', 'АФ №35', '35']}, {'Торговый зал ШФ №5': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №5.xlsx', 'ШФ №5', '5']}, {'Торговый зал ТФ №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ТФ №2.xlsx', 'ТФ №2', '2']}, {'Торговый зал АСФ №67': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №67.xlsx', 'АСФ №67', '67']}, {'Торговый зал АФ №49': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №49.xlsx', 'АФ №49', '49']}, {'Торговый зал АСФ №12': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №12.xlsx', 'АСФ №12', '12']}, {'Торговый зал ТФ №3': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ТФ №3.xlsx', 'ТФ №3', '3']}, {'Торговый зал АСФ №72': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №72.xlsx', 'АСФ №72', '72']}, {'Торговый зал ТЗФ №3': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ТЗФ №3.xlsx', 'ТЗФ №3', '3']}, {'Торговый зал АСФ №64': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №64.xlsx', 'АСФ №64', '64']}, {'Торговый зал АФ №59': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №59.xlsx', 'АФ №59', '59']}, {'Торговый зал АСФ №24': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №24.xlsx', 'АСФ №24', '24']}, {'Торговый зал АСФ №28': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №28.xlsx', 'АСФ №28', '28']}, {'Торговый зал АФ №51': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №51.xlsx', 'АФ №51', '51']}, {'Торговый зал АФ №64': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №64.xlsx', 'АФ №64', '64']}, {'Торговый зал ШФ №25': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №25.xlsx', 'ШФ №25', '25']}, {'Торговый зал ШФ №20': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №20.xlsx', 'ШФ №20', '20']}, {'Торговый зал АСФ №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №1.xlsx', 'АСФ №1', '1']}, {'Торговый зал АФ №33': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №33.xlsx', 'АФ №33', '33']}, {'Торговый зал ППФ №7': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №7.xlsx', 'ППФ №7', '7']}, {'Торговый зал УКФ №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал УКФ №1.xlsx', 'УКФ №1', '1']}, {'Торговый зал_ОПТ ШФ №35': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал_ОПТ ШФ №35.xlsx', 'Торговый зал_ОПТ ШФ №35', '35']}, {'Торговый зал АСФ №18': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №18.xlsx', 'АСФ №18', '18']}, {'Торговый зал АСФ №53': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №53.xlsx', 'АСФ №53', '53']}, {'Торговый зал АФ №58': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №58.xlsx', 'АФ №58', '58']}, {'Торговый зал ШФ №14': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №14.xlsx', 'ШФ №14', '14']}, {'Торговый зал КФ №5': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал КФ №5.xlsx', 'КФ №5', '5']}, {'Торговый зал АСФ №17': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №17.xlsx', 'АСФ №17', '17']}, {'Торговый зал АСФ №63': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №63.xlsx', 'АСФ №63', '63']}, {'Торговый зал АСФ №44': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №44.xlsx', 'АСФ №44', '44']}, {'Торговый зал КФ №6': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал КФ №6.xlsx', 'КФ №6', '6']}, {'Торговый зал АФ №47': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №47.xlsx', 'АФ №47', '47']}, {'Торговый зал АСФ №45': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №45.xlsx', 'АСФ №45', '45']}, {'Торговый зал АСФ №35': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №35.xlsx', 'АСФ №35', '35']}, {'Торговый зал АФ №77': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №77.xlsx', 'АФ №77', '77']}, {'Торговый зал АФ №17': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №17.xlsx', 'АФ №17', '17']}, {'Торговый зал АФ №75': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №75.xlsx', 'АФ №75', '75']}, {'Торговый зал ШФ №6': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №6.xlsx', 'ШФ №6', '6']}, {'Торговый зал КЗФ №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал КЗФ №2.xlsx', 'КЗФ №2', '2']}, {'Торговый зал ШФ №30': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №30.xlsx', 'ШФ №30', '30']}, {'Торговый зал КЗФ №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал КЗФ №1.xlsx', 'КЗФ №1', '1']}, {'Торговый зал АФ №24': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №24.xlsx', 'АФ №24', '24']}, {'Торговый зал АСФ №31': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №31.xlsx', 'АСФ №31', '31']}, {'Торговый зал АФ №78': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №78.xlsx', 'АФ №78', '78']}, {'Торговый зал ППФ №9': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №9.xlsx', 'ППФ №9', '9']}, {'Торговый зал АСФ №52': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №52.xlsx', 'АСФ №52', '52']}, {'Торговый зал АФ №61': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №61.xlsx', 'АФ №61', '61']}, {'Торговый зал АФ №9': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №9.xlsx', 'АФ №9', '9']}, {'Торговый зал АСФ №69': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №69.xlsx', 'АСФ №69', '69']}, {'Торговый зал ФКС №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ФКС №1.xlsx', 'ФКС №1', '1']}, {'Торговый зал АФ №14': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №14.xlsx', 'АФ №14', '14']}, {'Торговый зал ШФ №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №1.xlsx', 'ШФ №1', '1']}, {'Торговый зал АСФ №75': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №75.xlsx', 'АСФ №75', '75']}, {'Торговый зал АФ №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №1.xlsx', 'АФ №1', '1']}, {'Торговый зал ШФ №9': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №9.xlsx', 'ШФ №9', '9']}, {'Торговый зал ППФ №6': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №6.xlsx', 'ППФ №6', '6']}, {'Торговый зал АСФ №33': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №33.xlsx', 'АСФ №33', '33']}, {'Торговый зал ТКФ №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ТКФ №1.xlsx', 'ТКФ №1', '1']}, {'Торговый зал ШФ №19': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №19.xlsx', 'ШФ №19', '19']}, {'Торговый зал АСФ №80': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №80.xlsx', 'АСФ №80', '80']}, {'Торговый зал АФ №4': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №4.xlsx', 'АФ №4', '4']}, {'Торговый зал ППФ №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №2.xlsx', 'ППФ №2', '2']}, {'Торговый зал АФ №41': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №41.xlsx', 'АФ №41', '41']}, {'Торговый зал АФ №38': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №38.xlsx', 'АФ №38', '38']}, {'Торговый зал АФ №21': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №21.xlsx', 'АФ №21', '21']}, {'Торговый зал АФ №34': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №34.xlsx', 'АФ №34', '34']}, {'Торговый зал ППФ №21': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №21.xlsx', 'ППФ №21', '21']}, {'Торговый зал ШФ №3': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №3.xlsx', 'ШФ №3', '3']}, {'Торговый зал АСФ №30': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №30.xlsx', 'АСФ №30', '30']}, {'Торговый зал АСФ №76': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №76.xlsx', 'АСФ №76', '76']}, {'Торговый зал АФ №25': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №25.xlsx', 'АФ №25', '25']}, {'Торговый зал АФ №76': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №76.xlsx', 'АФ №76', '76']}, {'Торговый зал АСФ №54': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №54.xlsx', 'АСФ №54', '54']}, {'Торговый зал ТКФ №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ТКФ №2.xlsx', 'ТКФ №2', '2']}, {'Торговый зал ТЗФ №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ТЗФ №2.xlsx', 'ТЗФ №2', '2']}, {'Торговый зал ШФ №17': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №17.xlsx', 'ШФ №17', '17']}, {'Торговый зал АСФ №51': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №51.xlsx', 'АСФ №51', '51']}, {'Торговый зал КФ №7': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал КФ №7.xlsx', 'КФ №7', '7']}, {'Торговый зал АСФ №10': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №10.xlsx', 'АСФ №10', '10']}, {'Торговый зал АСФ №50': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №50.xlsx', 'АСФ №50', '50']}, {'Торговый зал АСФ №4': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №4.xlsx', 'АСФ №4', '4']}, {'Торговый зал АСФ №25': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №25.xlsx', 'АСФ №25', '25']}, {'Торговый зал АФ №52': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №52.xlsx', 'АФ №52', '52']}, {'Торговый зал АСФ №23': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №23.xlsx', 'АСФ №23', '23']}, {'Торговый зал ППФ №15': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №15.xlsx', 'ППФ №15', '15']}, {'Торговый зал АСФ №34': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №34.xlsx', 'АСФ №34', '34']}, {'Торговый зал ППФ №4': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №4.xlsx', 'ППФ №4', '4']}, {'Торговый зал СТМ 1АФ': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал СТМ 1АФ.xlsx', 'СТМ 1АФ', '1']}, {'Торговый зал КФ №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал КФ №2.xlsx', 'КФ №2', '2']}, {'Торговый зал ШФ №7': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №7.xlsx', 'ШФ №7', '7']}, {'Торговый зал ППФ №13': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №13.xlsx', 'ППФ №13', '13']}, {'Торговый зал АФ №81': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №81.xlsx', 'АФ №81', '81']}, {'Торговый зал АФ №32': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №32.xlsx', 'АФ №32', '32']}, {'Торговый зал АФ №45': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №45.xlsx', 'АФ №45', '45']}, {'Торговый зал АСФ №77': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №77.xlsx', 'АСФ №77', '77']}, {'Торговый зал ППФ №20': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №20.xlsx', 'ППФ №20', '20']}, {'Торговый зал АСФ №59': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №59.xlsx', 'АСФ №59', '59']}, {'Торговый зал ППФ №16': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №16.xlsx', 'ППФ №16', '16']}, {'Торговый зал ППФ №18': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №18.xlsx', 'ППФ №18', '18']}, {'Торговый зал АФ №10': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №10.xlsx', 'АФ №10', '10']}, {'Торговый_зал АФ №55 ОПТ': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый_зал АФ №55 ОПТ.xlsx', 'Торговый_зал АФ №55 ОПТ', '55']}, {'Торговый зал АФ №12': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №12.xlsx', 'АФ №12', '12']}, {'Торговый зал АФ №39': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №39.xlsx', 'АФ №39', '39']}, {'Торговый зал ШФ №18': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №18.xlsx', 'ШФ №18', '18']}, {'Торговый зал АФ №44': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №44.xlsx', 'АФ №44', '44']}, {'Торговый зал СТМ 6ШФ BAIS': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал СТМ 6ШФ BAIS.xlsx', 'СТМ 6ШФ BAIS', '6']}, {'Торговый зал ФКС №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ФКС №2.xlsx', 'ФКС №2', '2']}, {'Торговый зал АФ №23': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №23.xlsx', 'АФ №23', '23']}, {'Торговый зал АФ №43': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №43.xlsx', 'АФ №43', '43']}, {'Торговый зал АСФ №36': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №36.xlsx', 'АСФ №36', '36']}, {'Торговый зал АСФ №61': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №61.xlsx', 'АСФ №61', '61']}, {'Торговый зал АФ №48': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №48.xlsx', 'АФ №48', '48']}, {'Торговый зал АСФ №71': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №71.xlsx', 'АСФ №71', '71']}, {'Торговый зал ППФ №14': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №14.xlsx', 'ППФ №14', '14']}, {'Торговый зал АСФ №56': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №56.xlsx', 'АСФ №56', '56']}, {'Торговый зал АФ №70': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №70.xlsx', 'АФ №70', '70']}, {'Торговый зал АСФ №65': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №65.xlsx', 'АСФ №65', '65']}, {'Торговый зал АФ №68': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №68.xlsx', 'АФ №68', '68']}, {'Торговый зал ТФ №1': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ТФ №1.xlsx', 'ТФ №1', '1']}, {'Торговый зал АФ №36': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №36.xlsx', 'АФ №36', '36']}, {'Торговый зал АФ №84': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №84.xlsx', 'АФ №84', '84']}, {'Торговый зал ШФ №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №2.xlsx', 'ШФ №2', '2']}, {'Торговый зал АФ №15': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №15.xlsx', 'АФ №15', '15']}, {'Торговый зал ШФ №10': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №10.xlsx', 'ШФ №10', '10']}, {'Торговый зал АСФ №66': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №66.xlsx', 'АСФ №66', '66']}, {'Торговый зал ППФ №17': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №17.xlsx', 'ППФ №17', '17']}, {'Торговый зал АСФ №40': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №40.xlsx', 'АСФ №40', '40']}, {'Торговый зал КФ №4': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал КФ №4.xlsx', 'КФ №4', '4']}, {'Торговый зал ШФ №32': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №32.xlsx', 'ШФ №32', '32']}, {'Торговый зал ШФ №34': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №34.xlsx', 'ШФ №34', '34']}, {'Торговый зал АСФ №41': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №41.xlsx', 'АСФ №41', '41']}, {'Торговый зал АФ №80': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №80.xlsx', 'АФ №80', '80']}, {'Торговый зал АФ №72': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №72.xlsx', 'АФ №72', '72']}, {'Торговый зал ППФ №22': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №22.xlsx', 'ППФ №22', '22']}, {'Торговый зал АФ №19': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №19.xlsx', 'АФ №19', '19']}, {'Торговый зал ППФ №12': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №12.xlsx', 'ППФ №12', '12']}, {'Торговый зал АФ №56': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №56.xlsx', 'АФ №56', '56']}, {'Торговый зал АФ №37': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №37.xlsx', 'АФ №37', '37']}, {'Торговый зал АСФ №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №2.xlsx', 'АСФ №2', '2']}, {'Торговый зал АСФ №5': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №5.xlsx', 'АСФ №5', '5']}, {'Торговый зал ШФ №15': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №15.xlsx', 'ШФ №15', '15']}, {'Торговый зал АСФ №6': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №6.xlsx', 'АСФ №6', '6']}, {'Торговый зал АФ №54': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №54.xlsx', 'АФ №54', '54']}, {'Торговый зал АСФ №8': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №8.xlsx', 'АСФ №8', '8']}, {'Торговый зал АФ №16': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №16.xlsx', 'АФ №16', '16']}, {'Торговый зал АФ №46': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №46.xlsx', 'АФ №46', '46']}, {'Торговый зал АСФ №13': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №13.xlsx', 'АСФ №13', '13']}, {'Торговый зал АСФ №57': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №57.xlsx', 'АСФ №57', '57']}, {'Торговый зал АСФ №21': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №21.xlsx', 'АСФ №21', '21']}, {'Торговый зал АСФ №83': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №83.xlsx', 'АСФ №83', '83']}, {'Торговый зал АСФ №7': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №7.xlsx', 'АСФ №7', '7']}, {'Торговый зал УКФ №2': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал УКФ №2.xlsx', 'УКФ №2', '2']}, {'Торговый зал АСФ №37': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №37.xlsx', 'АСФ №37', '37']}, {'Торговый зал КФ №3': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал КФ №3.xlsx', 'КФ №3', '3']}, {'Торговый зал АФ №20': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №20.xlsx', 'АФ №20', '20']}, {'Торговый зал АСФ №9': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №9.xlsx', 'АСФ №9', '9']}, {'Торговый зал ППФ №5': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №5.xlsx', 'ППФ №5', '5']}, {'Торговый зал АФ №28': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №28.xlsx', 'АФ №28', '28']}, {'Торговый зал ШФ №4': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №4.xlsx', 'ШФ №4', '4']}, {'Торговый зал ППФ №10': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №10.xlsx', 'ППФ №10', '10']}, {'Торговый зал АФ №62': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №62.xlsx', 'АФ №62', '62']}, {'Торговый зал АФ №30': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №30.xlsx', 'АФ №30', '30']}, {'Торговый зал ШФ №27': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №27.xlsx', 'ШФ №27', '27']}, {'Торговый зал ШФ №8': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ШФ №8.xlsx', 'ШФ №8', '8']}, {'Торговый зал АФ №57': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №57.xlsx', 'АФ №57', '57']}, {'Торговый зал АСФ №46': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №46.xlsx', 'АСФ №46', '46']}, {'Торговый зал АФ №86': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АФ №86.xlsx', 'АФ №86', '86']}, {'Торговый зал ППФ №8': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №8.xlsx', 'ППФ №8', '8']}, {'Торговый зал АСФ №15': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №15.xlsx', 'АСФ №15', '15']}, {'Торговый зал АСФ №3': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №3.xlsx', 'АСФ №3', '3']}, {'Торговый зал АСФ №29': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №29.xlsx', 'АСФ №29', '29']}, {'Торговый зал АСФ №58': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №58.xlsx', 'АСФ №58', '58']}, {'Торговый зал ППФ №23': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал ППФ №23.xlsx', 'ППФ №23', '23']}, {'Торговый зал АСФ №62': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №62.xlsx', 'АСФ №62', '62']}, {'Торговый зал АСФ №82': ['\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-stat-1t\\Output\\Выгрузка 1Т из 1С\\Торговый зал АСФ №82.xlsx', 'АСФ №82', '82']}]
shuffle(a)
print(a)