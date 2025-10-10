#include <iostream>
#include <string>
#include <chrono>
#include <thread>
#include <cctype>

// libtorrent
#include <libtorrent/session.hpp>
#include <libtorrent/session_params.hpp>
#include <libtorrent/torrent_info.hpp>
#include <libtorrent/add_torrent_params.hpp>
#include <libtorrent/magnet_uri.hpp>
#include <libtorrent/alert_types.hpp>
#include <libtorrent/read_resume_data.hpp>
#include <libtorrent/write_resume_data.hpp>
#include <libtorrent/bencode.hpp>

// Windows specific for Unicode console (guarded)
#ifdef _WIN32
#  include <windows.h>
#endif

using namespace std::chrono_literals;

static bool isMagnetUri(const std::string &arg)
{
    // 한글 주석: magnet 링크인지 간단히 확인
    if (arg.size() < 7) return false;
    std::string prefix = arg.substr(0, 7);
    for (auto &ch : prefix) ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
    return prefix == "magnet:";
}

static std::string humanBytes(long long bytes)
{
    // 한글 주석: 바이트 단위를 사람이 읽기 쉽게 변환
    const char* units[] = {"B", "KB", "MB", "GB", "TB"};
    double count = static_cast<double>(bytes);
    int i = 0;
    while (count >= 1024.0 && i < 4) { count /= 1024.0; ++i; }
    char buf[64];
    std::snprintf(buf, sizeof(buf), "%.2f %s", count, units[i]);
    return std::string(buf);
}

static std::string formatEta(double seconds)
{
    // 한글 주석: ETA를 H:MM:SS 형식으로 변환
    if (seconds <= 0 || !std::isfinite(seconds)) return "--:--:--";
    auto total = static_cast<long long>(seconds + 0.5);
    long long h = total / 3600;
    long long m = (total % 3600) / 60;
    long long s = total % 60;
    char buf[32];
    if (h > 99) h = 99; // 너무 큰 값 방지
    std::snprintf(buf, sizeof(buf), "%02lld:%02lld:%02lld", h, m, s);
    return std::string(buf);
}

int main(int argc, char* argv[])
{
#ifdef _WIN32
    // 한글 주석: Windows 콘솔 유니코드 설정 (가능한 경우)
    SetConsoleOutputCP(CP_UTF8);
#endif

    if (argc < 2) {
        std::cerr << "Usage: libtorrent_downloader <.torrent|magnet> [save_path]\n";
        return 1;
    }

    std::string source = argv[1];
    // 한글 주석: 저장 경로는 인자가 없으면 현재 디렉토리 사용
    std::string savePath = ".";
    if (argc >= 3 && argv[2] != nullptr && std::strlen(argv[2]) > 0) {
        savePath = argv[2];
    }

    try {
        // 한글 주석: 세션 파라미터 설정 (업로드 억제)
        lt::settings_pack settings;
        settings.set_int(lt::settings_pack::alert_mask, lt::alert_category::status | lt::alert_category::error | lt::alert_category::storage);
        settings.set_int(lt::settings_pack::unchoke_slots_limit, 0);
        settings.set_int(lt::settings_pack::upload_rate_limit, 1); // 1 byte/s 사실상 업로드 차단
        settings.set_bool(lt::settings_pack::stop_when_ready, true); // 완료 즉시 정지 (시딩 안함)
        settings.set_bool(lt::settings_pack::enable_upnp, false);
        settings.set_bool(lt::settings_pack::enable_natpmp, false);
        settings.set_bool(lt::settings_pack::enable_lsd, false);
        settings.set_bool(lt::settings_pack::announce_to_all_trackers, true);
        settings.set_bool(lt::settings_pack::announce_to_all_tiers, true);
        settings.set_int(lt::settings_pack::connections_limit, 200);
        settings.set_int(lt::settings_pack::active_seeds, 0);
        settings.set_int(lt::settings_pack::active_limit, 1);

        lt::session ses(settings);

        lt::add_torrent_params atp;
        atp.save_path = savePath;
        atp.flags |= lt::torrent_flags::stop_when_ready;
        atp.flags &= ~lt::torrent_flags::upload_mode; // upload_mode는 다른 의미로 느릴 수 있어 해제
        atp.flags |= lt::torrent_flags::auto_managed;
        atp.flags |= lt::torrent_flags::paused; // 추가 직후 활성화는 세션에 맡김

        if (isMagnetUri(source)) {
            // 한글 주석: 마그넷 링크 파싱
            lt::error_code ec;
            lt::parse_magnet_uri(source, atp, ec);
            if (ec) {
                std::cerr << "Failed to parse magnet: " << ec.message() << "\n";
                return 2;
            }
        } else {
            // 한글 주석: 토렌트 파일 로드
            lt::error_code ec;
            lt::torrent_info ti(source, ec);
            if (ec || !ti.is_valid()) {
                std::cerr << "Failed to load .torrent: " << (ec ? ec.message() : "invalid torrent") << "\n";
                return 3;
            }
            atp.ti = std::make_shared<lt::torrent_info>(std::move(ti));
        }

        lt::torrent_handle th = ses.add_torrent(atp);
        th.resume(); // 시작

        bool done = false;
        auto lastPrint = std::chrono::steady_clock::now();

        while (!done) {
            // 한글 주석: 알럿 처리
            std::vector<lt::alert*> alerts;
            ses.pop_alerts(&alerts);
            for (lt::alert* a : alerts) {
                if (auto* st = lt::alert_cast<lt::state_update_alert>(a)) {
                    for (auto const& s : st->status) {
                        if (s.handle == th) {
                            if (s.is_seeding || s.is_finished || s.errc) {
                                done = true;
                            }
                        }
                    }
                } else if (auto* f = lt::alert_cast<lt::torrent_finished_alert>(a)) {
                    if (f->handle == th) done = true;
                } else if (auto* e = lt::alert_cast<lt::torrent_error_alert>(a)) {
                    std::cerr << "Torrent error: " << e->error.message() << "\n";
                    return 4;
                }
            }

            // 상태 업데이트 요청 (500ms마다)
            auto now = std::chrono::steady_clock::now();
            if (now - lastPrint >= 500ms) {
                lt::torrent_status st = th.status(lt::torrent_handle::query_save_path | lt::torrent_handle::query_name);

                float progress = st.progress_ppm / 10000.0f; // 0..100
                long long totalWanted = st.total_wanted;
                long long totalWantedDone = st.total_wanted_done;
                int64_t downRate = st.download_rate;
                int64_t upRate = st.upload_rate;

                double remaining = std::max<int64_t>(0, totalWanted - totalWantedDone);
                double etaSec = (downRate > 0) ? (remaining / static_cast<double>(downRate)) : std::numeric_limits<double>::infinity();

                std::cout << "[" << (st.is_finished ? "FIN" : "DL ") << "] "
                          << (progress) << "%  "
                          << humanBytes(totalWantedDone) << "/" << humanBytes(totalWanted)
                          << "  D:" << humanBytes(downRate) << "/s"
                          << "  U:" << humanBytes(upRate) << "/s"
                          << "  ETA " << formatEta(etaSec)
                          << "\r" << std::flush;

                lastPrint = now;

                // 완료 시 루프 종료
                if (st.is_finished || st.is_seeding) {
                    done = true;
                }

                // 다음 업데이트 위해 상태 갱신 요청
                ses.post_torrent_updates();
            }

            std::this_thread::sleep_for(50ms);
        }

        // 완료 메시지 및 정지
        lt::torrent_status st = th.status(lt::torrent_handle::query_save_path | lt::torrent_handle::query_name);
        std::cout << "\nCompleted: " << st.name << " -> " << st.save_path << "\n";
        th.pause();
        ses.pause();
    }
    catch (const std::exception& ex) {
        std::cerr << "Fatal error: " << ex.what() << "\n";
        return 10;
    }

    return 0;
}
