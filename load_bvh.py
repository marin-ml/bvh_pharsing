import re
import xlwt


def identifier(scanner, token):
    return "IDENT", token


def operator(scanner, token):
    return "OPERATOR", token


def digit(scanner, token):
    return "DIGIT", token


def open_brace(scanner, token):
    return "OPEN_BRACE", token


def close_brace(scanner, token):
    return "CLOSE_BRACE", token


def new_bone(parent, name):
    d_bone = {"parent": parent, "channels": [], "offsets": []}
    return d_bone


def push_bone_context(d_name):
    global bone_context
    bone_context.append(d_name)


def get_bone_context():
    global bone_context
    return bone_context[len(bone_context)-1]


def pop_bone_context():
    global bone_context
    bone_context = bone_context[:-1]
    return bone_context[len(bone_context)-1]


def read_offset(bvh, token_index):
    if bvh[token_index] != ("IDENT", "OFFSET"):
        return None, None

    token_index += 1
    offsets = [0.0] * 3
    for i in range(0, 3):
        offsets[i] = float(bvh[token_index][1])
        token_index += 1
    return offsets, token_index


def read_channels(bvh, token_index):
    if bvh[token_index] != ("IDENT", "CHANNELS"):
        return None, None

    token_index += 1
    channel_count = int(bvh[token_index][1])
    token_index += 1
    channels = [""] * channel_count
    for i in range(0, channel_count):
        channels[i] = bvh[token_index][1]
        token_index += 1
    return channels, token_index


def parse_joint(bvh, token_index):
    end_site = False
    joint_id = bvh[token_index][1]
    token_index += 1
    joint_name = bvh[token_index][1]
    token_index += 1

    if joint_id == "End":
        joint_name = get_bone_context() + "_Nub"
        end_site = True

    joint = new_bone(get_bone_context(), joint_name)

    if bvh[token_index][0] != "OPEN_BRACE":
        print "Was expecting brace, got ", bvh[token_index]
        return None

    token_index += 1
    offsets, token_index = read_offset(bvh, token_index)
    joint["offsets"] = offsets

    if not end_site:
        channels, token_index = read_channels(bvh, token_index)
        joint["channels"] = channels
        for channel in channels:
            motion_channels.append((joint_name, channel))

    skeleton[joint_name] = joint

    while ((bvh[token_index][0] == "IDENT") and (bvh[token_index][1] == "JOINT")) or \
            ((bvh[token_index][0] == "IDENT") and (bvh[token_index][1] == "End")):
        push_bone_context(joint_name)
        token_index = parse_joint(bvh, token_index)
        pop_bone_context()

    if (bvh[token_index][0]) == "CLOSE_BRACE":
        return token_index + 1

    print "Unexpected token ", bvh[token_index]


def parse_hierarchy(bvh):
    global current_token
    current_token = 0

    if bvh[current_token] != ("IDENT", "HIERARCHY"):
        return None

    current_token += 1

    if bvh[current_token] != ("IDENT", "ROOT"):
        return None

    current_token += 1

    if bvh[current_token][0] != "IDENT":
        return None

    root_name = bvh[current_token][1]
    root_bone = new_bone(None, root_name)

    current_token += 1

    if bvh[current_token][0] != "OPEN_BRACE":
        return None

    current_token += 1

    offsets, current_token = read_offset(bvh, current_token)
    channels, current_token = read_channels(bvh, current_token)
    root_bone["offsets"] = offsets
    root_bone["channels"] = channels
    skeleton[root_name] = root_bone
    push_bone_context(root_name)
    print "Root ", root_bone

    while bvh[current_token][1] == "JOINT":
        current_token = parse_joint(bvh, current_token)


def parse_motion(bvh):
    global current_token

    if bvh[current_token][0] != "IDENT":
        print "Unexpected text"
        return None

    if bvh[current_token][1] != "MOTION":
        print "No motion section"
        return None

    current_token += 1

    if bvh[current_token][1] != "Frames":
        return None

    current_token += 1
    frame_count = int(bvh[current_token][1])
    current_token += 1

    if bvh[current_token][1] != "Frame":
        return None

    current_token += 1

    if bvh[current_token][1] != "Time":
        return None

    current_token += 1
    frame_rate = float(bvh[current_token][1])
    frame_time = 0.0
    motions = [()] * frame_count

    frame_time_list = []
    value_list = []
    for i in range(0, frame_count):
        # print "Parsing frame ", i
        channel_values = []
        for channel in motion_channels:
            channel_values.append((channel[0], channel[1], float(bvh[current_token][1])))
            current_token += 1
        motions[i] = (frame_time, channel_values)
        # print motions[i]
        frame_time += frame_rate
        frame_time_list.append(frame_time)
        value_list.append(channel_values)

    return frame_time_list, value_list


def convert_list3(d):
    if d:
        return d
    else:
        return ['', '', '']


def convert_bvh(file_bvh):

    global current_token

    """ -------------- load the bvh file ----------- """
    bvh_file = open(file_bvh + '.bvh', "r")
    bvh = bvh_file.read()
    bvh_file.close()
    tokens, remainder = scanner.scan(bvh)
    parse_hierarchy(tokens)

    """ ---------------- parsing bvh file ----------- """
    set1 = [set1_title]
    for (name, bone) in skeleton.iteritems():
        set1.append([name] + convert_list3(bone["channels"][:3]) + bone["offsets"])

    """ ----------------- write to excel ------------ """
    book_out = xlwt.Workbook(encoding="utf-8")
    sheet1 = book_out.add_sheet("sheet1")
    sheet2 = book_out.add_sheet("sheet2")

    current_token += 1
    frame_time_set, value_set = parse_motion(tokens)

    for i in xrange(set1.__len__()):                # write to first sheet
        for j in xrange(set1[i].__len__()):
            sheet1.write(i, j, set1[i][j])

    sheet2.write(0, 0, 'Frame time')                # write the second sheet header
    sheet2.write(0, 1, 'Channel1')
    sheet2.write(1, 1, 'Channel2')

    for j in xrange(value_set[0].__len__()):
        sheet2.write(0, j + 2, value_set[0][j][0])
        sheet2.write(1, j + 2, value_set[0][j][1])

    for i in xrange(frame_time_set.__len__()):      # write the second sheet
        sheet2.write(i + 2, 0, frame_time_set[i])
        for j in xrange(value_set[i].__len__()):
            sheet2.write(i + 2, j + 2, value_set[i][j][2])

    book_out.save(bvh_file_name + '.xls')


current_token = 0
skeleton = {}
bone_context = []
motion_channels = []
motions = []
reserved = ["HIERARCHY", "ROOT", "OFFSET", "CHANNELS", "MOTION"]
channel_names = ["Xposition", "Yposition", "Zposition",  "Zrotation", "Xrotation",  "Yrotation"]
set1_title = ['Bone', 'Channel1', 'Channel2', 'Channel3', 'Offset1', 'Offset2', 'Offset3']

scanner = re.Scanner([
    (r"[a-zA-Z_]\w*", identifier),
    (r"-*[0-9]+(\.[0-9]+)?", digit),
    (r"}", close_brace),
    (r"{", open_brace),
    (r":", None),
    (r"\s+", None),
    ])

if __name__ == "__main__":

    bvh_file_name = "2017-02-23_20-33-44"
    # bvh_file_name = "2017-04-01_18-35-21"
    # bvh_file_name = "2017-04-01_20-56-08"
    convert_bvh(bvh_file_name)

    print "Complete successfully"
